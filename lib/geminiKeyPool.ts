import { GoogleGenerativeAI } from '@google/generative-ai'

const DEFAULT_MODEL_NAME = 'gemini-2.5-flash-lite'
const DEFAULT_MAX_ATTEMPTS = 3
const MAX_ALLOWED_ATTEMPTS = 10

let cachedKeys: string[] | null = null
let cachedSignature = ''
let roundRobinCursor = 0

type GeminiGenerateContentInput = Parameters<
    ReturnType<GoogleGenerativeAI['getGenerativeModel']>['generateContent']
>[0]

type GenerateGeminiContentWithRetryOptions = {
    contents: GeminiGenerateContentInput
    modelName?: string
    responseMimeType?: string
    maxAttempts?: number
}

type GenerateGeminiContentWithRetryResult = {
    text: string
    modelName: string
    keyIndex: number
    attempt: number
    totalAttempts: number
}

function getEnvSignature() {
    return `${process.env.GEMINI_API_KEYS || ''}::${process.env.GEMINI_API_KEY || ''}`
}

function normalizeAndDedupeKeys(keys: string[]) {
    return Array.from(new Set(keys.map((key) => key.trim()).filter(Boolean)))
}

function normalizeAndDedupeModels(models: string[]) {
    return Array.from(new Set(models.map((model) => model.trim()).filter(Boolean)))
}

function getGeminiApiKeys() {
    const currentSignature = getEnvSignature()

    if (cachedKeys && cachedSignature === currentSignature) {
        return cachedKeys
    }

    const keysFromPool = (process.env.GEMINI_API_KEYS || '').split(',')
    const singleKey = process.env.GEMINI_API_KEY || ''

    const normalizedKeys = normalizeAndDedupeKeys([
        ...keysFromPool,
        ...(singleKey ? [singleKey] : [])
    ])

    if (normalizedKeys.length === 0) {
        throw new Error('Gemini API key not configured. Set GEMINI_API_KEYS or GEMINI_API_KEY')
    }

    cachedKeys = normalizedKeys
    cachedSignature = currentSignature

    return normalizedKeys
}

function getNextGeminiApiKey(keys: string[]) {
    const keyIndex = roundRobinCursor % keys.length

    // Keep the counter bounded while preserving key distribution.
    roundRobinCursor = (roundRobinCursor + 1) % Number.MAX_SAFE_INTEGER

    return {
        keyIndex,
        apiKey: keys[keyIndex]
    }
}

function parseMaxAttempts(rawValue?: string, fallback = DEFAULT_MAX_ATTEMPTS) {
    if (!rawValue) {
        return fallback
    }

    const parsedValue = Number(rawValue)
    if (!Number.isFinite(parsedValue)) {
        return fallback
    }

    const normalized = Math.floor(parsedValue)
    if (normalized < 1) {
        return 1
    }

    return Math.min(normalized, MAX_ALLOWED_ATTEMPTS)
}

function getGeminiModelChain(primaryModelName: string) {
    const fallbackModels = normalizeAndDedupeModels((process.env.GEMINI_FALLBACK_MODELS || '').split(','))
    const modelChain = normalizeAndDedupeModels([primaryModelName, ...fallbackModels])

    if (modelChain.length < 2) {
        throw new Error(
            'Gemini fallback models not configured. Set GEMINI_FALLBACK_MODELS with at least one fallback model different from GEMINI_MODEL'
        )
    }

    return modelChain
}

function extractStatusCode(error: unknown) {
    if (!error || typeof error !== 'object') {
        return null
    }

    const candidateValues = [
        (error as { status?: unknown }).status,
        (error as { statusCode?: unknown }).statusCode,
        (error as { code?: unknown }).code,
        (error as { response?: { status?: unknown } }).response?.status,
        (error as { error?: { code?: unknown } }).error?.code
    ]

    for (const candidate of candidateValues) {
        const numericValue = typeof candidate === 'number'
            ? candidate
            : typeof candidate === 'string'
                ? Number(candidate)
                : NaN

        if (Number.isInteger(numericValue) && numericValue >= 100 && numericValue <= 599) {
            return numericValue
        }
    }

    return null
}

function getErrorMessage(error: unknown) {
    if (error instanceof Error && error.message) {
        return error.message
    }

    if (typeof error === 'string' && error) {
        return error
    }

    return 'Unknown Gemini error'
}

function isRetryableGeminiError(error: unknown) {
    const statusCode = extractStatusCode(error)

    if (statusCode !== null) {
        if (statusCode >= 500) {
            return true
        }

        if (statusCode === 401 || statusCode === 403 || statusCode === 408 || statusCode === 409 || statusCode === 429) {
            return true
        }
    }

    const message = getErrorMessage(error).toLowerCase()

    return /rate limit|quota|resource_exhausted|timeout|timed out|temporar|network|unavailable|econnreset|eai_again|etimedout/.test(message)
}

function maskApiKey(apiKey: string) {
    if (apiKey.length <= 4) {
        return '***'
    }

    return `***${apiKey.slice(-4)}`
}

function createGeminiPoolError(message: string, cause: unknown) {
    const error = new Error(message)
    ;(error as Error & { cause?: unknown }).cause = cause
    return error
}

export function getGeminiKeyPoolSize() {
    return getGeminiApiKeys().length
}

export async function generateGeminiContentWithRetry(
    options: GenerateGeminiContentWithRetryOptions
): Promise<GenerateGeminiContentWithRetryResult> {
    const apiKeys = getGeminiApiKeys()
    const primaryModelName = options.modelName || process.env.GEMINI_MODEL || DEFAULT_MODEL_NAME
    const modelChain = getGeminiModelChain(primaryModelName)
    const responseMimeType = options.responseMimeType || 'application/json'
    const maxAttempts = options.maxAttempts ?? parseMaxAttempts(process.env.GEMINI_MAX_ATTEMPTS)
    const totalAttempts = Math.max(1, maxAttempts)

    let lastError: unknown = null
    let lastModelName = primaryModelName

    for (let attempt = 1; attempt <= totalAttempts; attempt++) {
        const modelIndex = (attempt - 1) % modelChain.length
        const modelName = modelChain[modelIndex]
        const isFallbackModel = modelIndex > 0
        const { keyIndex, apiKey } = getNextGeminiApiKey(apiKeys)

        try {
            const genAI = new GoogleGenerativeAI(apiKey)
            const model = genAI.getGenerativeModel({
                model: modelName,
                generationConfig: {
                    responseMimeType
                }
            })

            const result = await model.generateContent(options.contents)
            const response = await result.response
            const text = response.text()

            return {
                text,
                modelName,
                keyIndex,
                attempt,
                totalAttempts
            }
        } catch (error) {
            lastError = error
            lastModelName = modelName

            const shouldRetry = attempt < totalAttempts && isRetryableGeminiError(error)
            const safeKey = maskApiKey(apiKey)
            const errorMessage = getErrorMessage(error)

            console.error(
                `[GeminiKeyPool] Attempt ${attempt}/${totalAttempts} failed on key #${keyIndex} (${safeKey}) with model "${modelName}"${isFallbackModel ? ` (fallback ${modelIndex}/${modelChain.length - 1})` : ' (primary)'}: ${errorMessage}`
            )

            if (!shouldRetry) {
                throw createGeminiPoolError(
                    `Gemini call failed after ${attempt} attempt(s). Last model "${modelName}". Model chain: ${modelChain.join(' -> ')}`,
                    error
                )
            }
        }
    }

    throw createGeminiPoolError(
        `Gemini call failed after ${totalAttempts} attempt(s). Last model "${lastModelName}". Model chain: ${modelChain.join(' -> ')}`,
        lastError
    )
}
