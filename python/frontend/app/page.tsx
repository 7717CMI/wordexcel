'use client'

import { useState, useRef, useEffect, useCallback } from 'react'
import FileUpload from './components/FileUpload'
import BulkFileUpload from './components/BulkFileUpload'
import DataReview from './components/DataReview'
import { ExtractedData } from './types'

type Step = 'upload' | 'review' | 'complete'
type Tab = 'single' | 'bulk'

interface FileResult {
  filename: string
  file_id: string
  market_name: string
  excel_filename: string
  excel_path: string
  success: boolean
  status: string
  error?: string
}

export default function Home() {
  const [activeTab, setActiveTab] = useState<Tab>('single')
  const [currentStep, setCurrentStep] = useState<Step>('upload')
  const [fileId, setFileId] = useState<string>('')
  const [extractedData, setExtractedData] = useState<ExtractedData | null>(null)
  const [confirmedData, setConfirmedData] = useState<ExtractedData | null>(null)
  const [error, setError] = useState<string>('')
  const [processingStep, setProcessingStep] = useState<string>('')
  const [_quickProcess, _setQuickProcess] = useState<boolean>(true)

  // Bulk processing states
  const [bulkResults, setBulkResults] = useState<FileResult[]>([])
  const [bulkProgress, setBulkProgress] = useState<{ completed: number; total: number }>({
    completed: 0,
    total: 0,
  })
  const [bulkErrors, setBulkErrors] = useState<FileResult[]>([])
  const [downloadedFiles, setDownloadedFiles] = useState<Set<string>>(new Set())
  const [isProcessing, setIsProcessing] = useState<boolean>(false)

  // Real-time progress states
  const [fileProgress, setFileProgress] = useState<
    Map<string, { status: string; message: string; error?: string }>
  >(new Map())
  const [currentProcessingFile, setCurrentProcessingFile] = useState<string>('')
  const [processingMessage, setProcessingMessage] = useState<string>('')

  // WebSocket connection
  const [wsConnected, setWsConnected] = useState<boolean>(false)
  const wsRef = useRef<WebSocket | null>(null)

  const downloadFile = useCallback(
    async (downloadUrl: string, filename: string) => {
      try {
        // Check if file has already been downloaded
        if (downloadedFiles.has(filename)) {
          console.log(`File already downloaded: ${filename}`)
          // Send confirmation anyway to prevent backend blocking
          if (wsRef.current && wsRef.current.readyState === WebSocket.OPEN) {
            wsRef.current.send(
              JSON.stringify({
                type: 'download_confirmed',
                file_id: filename,
              })
            )
            console.log(`Download confirmation sent for already downloaded file: ${filename}`)
          }
          return true
        }

        const downloadResponse = await fetch(downloadUrl)
        if (downloadResponse.ok) {
          const blob = await downloadResponse.blob()
          const url = window.URL.createObjectURL(blob)
          const a = document.createElement('a')
          a.href = url
          a.download = filename
          document.body.appendChild(a)
          a.click()
          window.URL.revokeObjectURL(url)
          document.body.removeChild(a)
          setDownloadedFiles((prev) => new Set(Array.from(prev).concat(filename)))
          console.log(`Auto-downloaded: ${filename}`)

          // Send download confirmation to backend
          if (wsRef.current && wsRef.current.readyState === WebSocket.OPEN) {
            wsRef.current.send(
              JSON.stringify({
                type: 'download_confirmed',
                file_id: filename,
              })
            )
            console.log(`Download confirmation sent for: ${filename}`)
          }

          return true
        } else {
          console.error(`Failed to download ${filename}`)
          return false
        }
      } catch (downloadError) {
        console.error(`Failed to download ${filename}:`, downloadError)
        return false
      }
    },
    [downloadedFiles, wsRef]
  )

  const handleWebSocketMessage = useCallback(
    (e: MessageEvent) => {
      const message = JSON.parse(e.data)
      console.log('WebSocket message received:', message)

      switch (message.type) {
        case 'bulk_start':
          setIsProcessing(true)
          setProcessingMessage(message.message)
          // Initialize file progress for all files - use STRING keys consistently
          const initialProgress = new Map()
          for (let i = 0; i < message.total_files; i++) {
            const key = String(i) // Always use string keys
            initialProgress.set(key, {
              status: 'pending',
              message: `File ${i + 1} waiting...`,
            })
            console.log(`Initialized file ${key} as pending`)
          }
          setFileProgress(initialProgress)
          setBulkResults([])
          setBulkErrors([])
          setBulkProgress({ completed: 0, total: message.total_files })
          console.log(
            `Bulk processing started for ${message.total_files} files, Map size: ${initialProgress.size}`
          )
          break

        case 'file_start':
          setCurrentProcessingFile(message.filename)
          setFileProgress((prev) => {
            const newMap = new Map(prev)
            const index = String(message.file_index) // Use string key
            console.log(
              `file_start - Index: ${index}, Type: ${typeof index}, Filename: ${message.filename}`
            )
            console.log(`Before update - Map has key ${index}:`, newMap.has(index))
            newMap.set(index, {
              status: 'processing',
              message: message.message || `Processing ${message.filename}...`,
            })
            console.log(`After update - Map size: ${newMap.size}, Keys:`, Array.from(newMap.keys()))
            return newMap
          })
          break

        case 'file_progress':
          setFileProgress((prev) => {
            const newMap = new Map(prev)
            const index = String(message.file_index) // Use string key
            newMap.set(index, {
              status: 'processing',
              message: message.message,
            })
            console.log(`file_progress - Index: ${index}, Map size: ${newMap.size}`)
            return newMap
          })
          break

        case 'file_completed':
          // Always update the file progress status
          setFileProgress((prev) => {
            const newMap = new Map(prev)
            const index = String(message.file_index) // Use string key
            console.log(
              `file_completed - Index: ${index}, Type: ${typeof index}, Filename: ${message.filename}`
            )
            console.log(`Before update - Map has key ${index}:`, newMap.has(index))
            newMap.set(index, {
              status: 'success',
              message: message.message || `${message.filename} completed successfully`,
            })
            console.log(`After update - Map size: ${newMap.size}, Keys:`, Array.from(newMap.keys()))
            console.log(
              `File statuses:`,
              Array.from(newMap.entries()).map(([k, v]) => `${k}: ${v.status}`)
            )
            return newMap
          })

          // Update bulk progress
          setBulkProgress((prev) => {
            const newProgress = { ...prev, completed: prev.completed + 1 }
            console.log(`Bulk progress: ${newProgress.completed}/${newProgress.total}`)
            return newProgress
          })

          // Check if this file has already been processed for download
          if (downloadedFiles.has(message.excel_filename)) {
            console.log(`File already downloaded: ${message.excel_filename}`)
            break
          }

          // Add to bulk results immediately
          const newResult: FileResult = {
            filename: message.filename,
            file_id: message.excel_filename, // Use excel_filename as file_id
            market_name: message.market_name,
            excel_filename: message.excel_filename,
            excel_path: message.download_url,
            success: true,
            status: 'completed',
          }
          setBulkResults((prev) => [...prev, newResult])

          // Auto-download the file immediately
          setTimeout(async () => {
            const downloadSuccess = await downloadFile(message.download_url, message.excel_filename)
            if (downloadSuccess) {
              console.log(`Successfully downloaded and confirmed: ${message.excel_filename}`)
            } else {
              console.error(`Failed to download: ${message.excel_filename}`)
              // Send failure confirmation to backend anyway to prevent blocking
              if (wsRef.current && wsRef.current.readyState === WebSocket.OPEN) {
                wsRef.current.send(
                  JSON.stringify({
                    type: 'download_confirmed',
                    file_id: message.excel_filename,
                  })
                )
              }
            }
          }, 500) // Small delay to prevent browser blocking
          break

        case 'file_failed':
          setFileProgress((prev) => {
            const newMap = new Map(prev)
            const index = String(message.file_index) // Use string key
            console.log(
              `file_failed - Index: ${index}, Type: ${typeof index}, Filename: ${message.filename}`
            )
            newMap.set(index, {
              status: 'failed',
              message: message.message || `Failed to process ${message.filename}`,
              error: message.error,
            })
            console.log(
              `After failed update - Map size: ${newMap.size}, Keys:`,
              Array.from(newMap.keys())
            )
            return newMap
          })
          setBulkProgress((prev) => ({ ...prev, completed: prev.completed + 1 }))

          // Add to bulk errors immediately
          const newError: FileResult = {
            filename: message.filename,
            file_id: '',
            market_name: 'Unknown Market',
            excel_filename: '',
            excel_path: '',
            success: false,
            status: 'failed',
            error: message.error,
          }
          setBulkErrors((prev) => [...prev, newError])
          break

        case 'bulk_complete':
          setIsProcessing(false)
          setCurrentProcessingFile('')
          setProcessingMessage(message.message || 'All files processed successfully!')
          console.log('Bulk processing complete:', message)

          // Ensure all pending files are marked as complete if not already
          setFileProgress((prev) => {
            const newMap = new Map(prev)
            Array.from(newMap.entries()).forEach(([key, value]) => {
              if (value.status === 'pending' || value.status === 'processing') {
                console.log(`Marking file ${key} as completed in bulk_complete`)
                newMap.set(key, {
                  ...value,
                  status: 'success',
                  message: value.message
                    .replace('waiting', 'completed')
                    .replace('Processing', 'Completed'),
                })
              }
            })
            return newMap
          })
          break

        case 'pong':
          console.log('Received pong from server')
          break

        default:
          console.log('Unknown message type:', message.type)
      }
    },
    [downloadFile, downloadedFiles]
  )

  // Initialize WebSocket connection
  useEffect(() => {
    let reconnectTimeout: NodeJS.Timeout | null = null
    let isConnecting = false
    let shouldReconnect = true

    const connectWebSocket = () => {
      // Prevent multiple simultaneous connections
      if (isConnecting || (wsRef.current && wsRef.current.readyState === WebSocket.CONNECTING)) {
        return
      }

      // Close existing connection if any
      if (wsRef.current) {
        wsRef.current.close()
      }

      isConnecting = true
      const wsProtocol = window.location.protocol === 'https:' ? 'wss:' : 'ws:'
      const wsHost = window.location.host
      const ws = new WebSocket(`${wsProtocol}//${wsHost}/ws`)

      ws.onopen = () => {
        console.log('WebSocket connected')
        setWsConnected(true)
        isConnecting = false
      }

      ws.onmessage = (event) => {
        console.log('Raw WebSocket message received:', event.data)
        try {
          handleWebSocketMessage(event)
        } catch (error) {
          console.error('Error handling WebSocket message:', error)
        }
      }

      ws.onclose = (event) => {
        console.log('WebSocket disconnected', event.code, event.reason)
        setWsConnected(false)
        isConnecting = false

        // Clear ping interval when connection closes
        if ((ws as any).pingInterval) {
          clearInterval((ws as any).pingInterval)
        }

        // Only reconnect if it wasn't a manual close (code 1000) and we should reconnect
        if (event.code !== 1000 && shouldReconnect) {
          console.log(
            `WebSocket closed with code ${event.code}, reason: ${event.reason}. Attempting to reconnect...`
          )
          // Clear any existing timeout
          if (reconnectTimeout) {
            clearTimeout(reconnectTimeout)
          }
          // Reconnect after 2 seconds (faster retry)
          reconnectTimeout = setTimeout(() => {
            if (shouldReconnect) {
              console.log('Attempting WebSocket reconnection...')
              connectWebSocket()
            }
          }, 2000)
        }
      }

      ws.onerror = (error) => {
        console.error('WebSocket error:', error)
        setWsConnected(false)
        isConnecting = false
      }

      // Add ping/pong to keep connection alive
      const pingInterval = setInterval(() => {
        if (ws.readyState === WebSocket.OPEN) {
          ws.send(JSON.stringify({ type: 'ping' }))
        }
      }, 30000) // Ping every 30 seconds

      // Store ping interval to clear it later
      ;(ws as any).pingInterval = pingInterval

      wsRef.current = ws
    }

    connectWebSocket()

    return () => {
      shouldReconnect = false

      // Clear reconnect timeout
      if (reconnectTimeout) {
        clearTimeout(reconnectTimeout)
      }

      // Close WebSocket connection
      if (wsRef.current) {
        // Clear ping interval if it exists
        if ((wsRef.current as any).pingInterval) {
          clearInterval((wsRef.current as any).pingInterval)
        }
        wsRef.current.close(1000, 'Component unmounting')
        wsRef.current = null
      }
    }
  }, []) // Remove dependency to prevent Fast Refresh reconnections

  const handleFileUploaded = async (uploadedFileId: string, quickProcess: boolean = true) => {
    setFileId(uploadedFileId)
    setError('')

    if (quickProcess) {
      await handleQuickProcess(uploadedFileId)
    } else {
      await handleManualProcess(uploadedFileId)
    }
  }

  const handleManualProcess = async (fileId: string) => {
    try {
      setProcessingStep('Extracting data from document...')

      // Step 1: Process document (extract data only)
      const processResponse = await fetch('/api/process', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({ fileId }),
      })

      if (!processResponse.ok) {
        const errorData = await processResponse.json()
        throw new Error(errorData.detail || 'Failed to process document')
      }

      const processResult = await processResponse.json()
      setExtractedData(processResult.data)
      setProcessingStep('')
      setCurrentStep('review')
    } catch (error) {
      console.error('Manual process error:', error)
      setError(
        `Failed to extract data: ${error instanceof Error ? error.message : 'Unknown error'}`
      )
      setProcessingStep('')
    }
  }

  const handleExcelGeneration = async (fileId: string, data: ExtractedData) => {
    try {
      setProcessingStep('Generating Excel file...')

      // Generate Excel
      const excelResponse = await fetch('/api/generate-excel', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          fileId,
          data,
        }),
      })

      if (!excelResponse.ok) {
        const errorData = await excelResponse.json()
        throw new Error(errorData.detail || 'Failed to generate Excel file')
      }

      // Download Excel file
      const blob = await excelResponse.blob()
      const marketName = data.market?.market_name || 'Market Data'
      const cleanMarketName =
        marketName.replace(/[<>:"/\\|?*\x00]/g, '').trim() || `Market Data ${Date.now()}`
      const filename = `${cleanMarketName}.xlsm`

      const url = window.URL.createObjectURL(blob)
      const a = document.createElement('a')
      a.href = url
      a.download = filename
      document.body.appendChild(a)
      a.click()
      window.URL.revokeObjectURL(url)
      document.body.removeChild(a)

      setCurrentStep('complete')
      setProcessingStep('')
    } catch (error) {
      console.error('Excel generation error:', error)
      setError(
        `Failed to generate Excel: ${error instanceof Error ? error.message : 'Unknown error'}`
      )
      setProcessingStep('')
    }
  }

  const handleQuickProcess = async (fileId: string) => {
    try {
      setProcessingStep('Extracting data from document...')

      // Step 1: Process document
      const processResponse = await fetch('/api/process', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({ fileId }),
      })

      if (!processResponse.ok) {
        const errorData = await processResponse.json()
        throw new Error(errorData.detail || 'Failed to process document')
      }

      const processResult = await processResponse.json()
      setExtractedData(processResult.data)
      setConfirmedData(processResult.data)
      setProcessingStep('Generating Excel file...')

      // Step 2: Generate Excel
      const excelResponse = await fetch('/api/generate-excel', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          fileId,
          data: processResult.data,
        }),
      })

      if (!excelResponse.ok) {
        const errorData = await excelResponse.json()
        throw new Error(errorData.detail || 'Failed to generate Excel file')
      }

      // Step 3: Download Excel file
      const blob = await excelResponse.blob()
      const marketName = processResult.data.market?.market_name || 'Market Data'
      const cleanMarketName =
        marketName.replace(/[<>:"/\\|?*\x00]/g, '').trim() || `Market Data ${Date.now()}`
      const filename = `${cleanMarketName}.xlsm`

      const url = window.URL.createObjectURL(blob)
      const a = document.createElement('a')
      a.href = url
      a.download = filename
      document.body.appendChild(a)
      a.click()
      window.URL.revokeObjectURL(url)
      document.body.removeChild(a)

      setCurrentStep('complete')
      setProcessingStep('')
    } catch (error) {
      console.error('Quick process error:', error)
      setError(
        `Failed to process file: ${error instanceof Error ? error.message : 'Unknown error'}`
      )
      setProcessingStep('')
    }
  }

  const handleFilesSelected = (files: File[]) => {
    // Clear file progress when new files are selected
    setFileProgress(new Map())
    setBulkResults([])
    setBulkErrors([])
    setDownloadedFiles(new Set())
    setBulkProgress({ completed: 0, total: files.length })
    setError('')
  }

  const handleBulkUploaded = async (files: File[]) => {
    try {
      setIsProcessing(true)
      setError('')
      setBulkResults([])
      setBulkErrors([])
      setDownloadedFiles(new Set())
      setBulkProgress({ completed: 0, total: files.length })
      setFileProgress(new Map()) // Reset file progress for new batch

      const formData = new FormData()
      files.forEach((file) => {
        formData.append('files', file)
      })

      console.log(`Starting bulk processing of ${files.length} files...`)

      // Manually initialize file progress for all files before API call
      const initialProgress = new Map()
      for (let i = 0; i < files.length; i++) {
        const key = String(i)
        initialProgress.set(key, {
          status: 'pending',
          message: `File ${i + 1} waiting...`,
        })
        console.log(`Pre-initialized file ${key} as pending`)
      }
      setFileProgress(initialProgress)
      setBulkProgress({ completed: 0, total: files.length })
      setIsProcessing(true)
      console.log(`Manual initialization: ${files.length} files, Map size: ${initialProgress.size}`)

      // Retry mechanism for connection issues
      let response
      let retryCount = 0
      const maxRetries = 2

      while (retryCount <= maxRetries) {
        try {
          response = await fetch('/api/independent-bulk-process', {
            method: 'POST',
            body: formData,
            // Add timeout for longer processing - 25 minutes
            signal: AbortSignal.timeout(1500000), // 25 minutes timeout
          })
          break // Success, exit retry loop
        } catch (fetchError) {
          retryCount++
          console.log(`Attempt ${retryCount} failed:`, fetchError.message)

          if (retryCount > maxRetries) {
            throw fetchError // Re-throw if all retries failed
          }

          // Wait before retrying
          await new Promise((resolve) => setTimeout(resolve, 2000))
          console.log(`Retrying... (${retryCount}/${maxRetries})`)
        }
      }

      console.log('Response received:')
      console.log('Status:', response.status)
      console.log('Status text:', response.statusText)
      console.log('Headers:', Object.fromEntries(response.headers.entries()))

      if (!response.ok) {
        const errorText = await response.text()
        console.error('Error response text:', errorText)
        let errorData
        try {
          errorData = JSON.parse(errorText)
        } catch (e) {
          console.error('Failed to parse error response as JSON:', e)
          throw new Error(`HTTP ${response.status}: ${response.statusText}`)
        }
        throw new Error(errorData.detail || 'Failed to start bulk processing')
      }

      const responseText = await response.text()
      console.log('Response text:', responseText)

      let result
      try {
        result = JSON.parse(responseText)
      } catch (e) {
        console.error('Failed to parse response as JSON:', e)
        console.error('Response text:', responseText)
        throw new Error('Backend returned invalid JSON response')
      }

      console.log('Parsed result:', result)

      // Don't process results here - WebSocket will handle individual file updates
      console.log(
        'Bulk processing started successfully. WebSocket will handle individual file updates.'
      )

      // Note: We don't set currentStep to 'complete' here because
      // WebSocket will handle the completion via 'bulk_complete' message
    } catch (error) {
      console.error('Bulk upload error:', error)
      console.error('Error details:', {
        name: error.name,
        message: error.message,
        stack: error.stack,
      })

      // Check if it's a connection reset error
      if (
        error.message.includes('socket hang up') ||
        error.message.includes('ECONNRESET') ||
        error.message.includes('connection reset')
      ) {
        setError(
          `Bulk processing failed: Connection was reset during processing. The backend may have completed successfully, but the connection was lost. Please check the temp folder for generated files.`
        )

        // Try to check if files were actually generated
        try {
          const checkResponse = await fetch('/api/check-generated-files', {
            method: 'POST',
            headers: {
              'Content-Type': 'application/json',
            },
            body: JSON.stringify({ filenames: files.map((f) => f.name) }),
          })

          if (checkResponse.ok) {
            const checkResult = await checkResponse.json()
            if (checkResult.files && checkResult.files.length > 0) {
              setError(
                `Connection was reset, but ${checkResult.files.length} files were successfully generated! Check the temp folder for: ${checkResult.files.join(', ')}`
              )
            }
          }
        } catch (checkError) {
          console.log('Could not check for generated files:', checkError)
        }
      }
      // Check if it's a timeout error
      else if (error.name === 'AbortError' || error.message.includes('timeout')) {
        setError(
          `Bulk processing failed: Request timed out after 25 minutes. Processing 3+ files may take longer. Please try again or process fewer files at once.`
        )
      }
      // Check if it's a JSON parsing error
      else if (
        error.message.includes('Unexpected token') ||
        error.message.includes('invalid JSON')
      ) {
        setError(
          `Bulk processing failed: Backend returned invalid response. Please check if the backend server is running correctly.`
        )
      } else {
        setError(
          `Bulk processing failed: ${error instanceof Error ? error.message : 'Unknown error'}`
        )
      }
    } finally {
      setIsProcessing(false)
    }
  }

  const handleIndividualDownload = async (fileResult: FileResult) => {
    try {
      const downloadResponse = await fetch(`/api/download/${fileResult.excel_filename}`)
      if (downloadResponse.ok) {
        const blob = await downloadResponse.blob()
        const url = window.URL.createObjectURL(blob)
        const a = document.createElement('a')
        a.href = url
        a.download = fileResult.excel_filename
        document.body.appendChild(a)
        a.click()
        window.URL.revokeObjectURL(url)
        document.body.removeChild(a)
        setDownloadedFiles((prev) => new Set(Array.from(prev).concat(fileResult.excel_filename)))
      } else {
        console.error(`Failed to download ${fileResult.excel_filename}`)
      }
    } catch (downloadError) {
      console.error(`Failed to download ${fileResult.excel_filename}:`, downloadError)
    }
  }

  const handleDownloadAll = async () => {
    try {
      // Download all successful files one by one
      for (const result of bulkResults) {
        if (!downloadedFiles.has(result.excel_filename)) {
          const downloadResponse = await fetch(`/api/download/${result.excel_filename}`)
          if (downloadResponse.ok) {
            const blob = await downloadResponse.blob()
            const url = window.URL.createObjectURL(blob)
            const a = document.createElement('a')
            a.href = url
            a.download = result.excel_filename
            document.body.appendChild(a)
            a.click()
            window.URL.revokeObjectURL(url)
            document.body.removeChild(a)

            // Add to downloaded files set
            setDownloadedFiles((prev) => new Set(Array.from(prev).concat(result.excel_filename)))

            // Small delay to prevent browser from blocking multiple downloads
            await new Promise((resolve) => setTimeout(resolve, 100))
          }
        }
      }
    } catch (error) {
      console.error('Failed to download all files:', error)
    }
  }

  const handleReset = () => {
    setCurrentStep('upload')
    setFileId('')
    setExtractedData(null)
    setConfirmedData(null)
    setError('')
    setProcessingStep('')
    setBulkResults([])
    setBulkProgress({ completed: 0, total: 0 })
    setBulkErrors([])
    setDownloadedFiles(new Set())
    setIsProcessing(false)
    setFileProgress(new Map())
    setCurrentProcessingFile('')
    setProcessingMessage('')
  }

  const renderContent = () => {
    if (activeTab === 'single') {
      switch (currentStep) {
        case 'upload':
          return (
            <div className="space-y-6">
              <div className="bg-white p-6 rounded-lg shadow-sm border">
                <h2 className="text-xl font-semibold text-gray-900 mb-4">Single File Processing</h2>
                <FileUpload onFileUploaded={handleFileUploaded} />
              </div>
            </div>
          )
        case 'review':
          return (
            <div className="space-y-6">
              <div className="bg-white p-6 rounded-lg shadow-sm border">
                <h2 className="text-xl font-semibold text-gray-900 mb-4">Review Extracted Data</h2>
                {extractedData && (
                  <DataReview
                    data={extractedData}
                    onDataConfirmed={async (confirmedData) => {
                      setConfirmedData(confirmedData)
                      await handleExcelGeneration(fileId, confirmedData)
                    }}
                    onBack={() => setCurrentStep('upload')}
                  />
                )}
              </div>
            </div>
          )
        case 'complete':
          return (
            <div className="space-y-6">
              <div className="bg-white p-6 rounded-lg shadow-sm border">
                <h2 className="text-xl font-semibold text-gray-900 mb-4">Processing Complete</h2>
                <div className="text-center">
                  <div className="text-green-600 text-6xl mb-4">✓</div>
                  <h3 className="text-lg font-medium text-gray-900 mb-2">
                    Excel file generated successfully!
                  </h3>
                  <p className="text-gray-600 mb-6">
                    Your Excel file has been downloaded with the market name:{' '}
                    <strong>{confirmedData?.market?.market_name}</strong>
                  </p>
                  <button
                    onClick={handleReset}
                    className="bg-blue-600 text-white px-6 py-2 rounded-md hover:bg-blue-700 transition-colors"
                  >
                    Process Another File
                  </button>
                </div>
              </div>
            </div>
          )
      }
    } else if (activeTab === 'bulk') {
      switch (currentStep) {
        case 'upload':
          return (
            <div className="space-y-6">
              <div className="bg-white p-6 rounded-lg shadow-sm border">
                <h2 className="text-xl font-semibold text-gray-900 mb-4">Bulk File Processing</h2>
                <div className="mb-4 p-4 bg-blue-50 border border-blue-200 rounded-md">
                  <div className="flex items-center">
                    <div
                      className={`w-3 h-3 rounded-full mr-3 ${wsConnected ? 'bg-green-500' : 'bg-red-500'}`}
                    ></div>
                    <span className="text-sm text-gray-700">
                      {wsConnected ? 'Ready for bulk processing' : 'Connecting to server...'}
                    </span>
                  </div>
                </div>
                <BulkFileUpload
                  onBulkUploaded={handleBulkUploaded}
                  isProcessing={isProcessing}
                  fileProgress={fileProgress}
                  onFilesSelected={handleFilesSelected}
                />
              </div>
            </div>
          )
        case 'complete':
          return (
            <div className="space-y-6">
              <div className="bg-white p-6 rounded-lg shadow-sm border">
                <h2 className="text-xl font-semibold text-gray-900 mb-4">
                  {isProcessing ? 'Processing Files...' : 'Bulk Processing Complete'}
                </h2>

                {/* Completion Status Banner */}
                {!isProcessing &&
                  bulkProgress.completed === bulkProgress.total &&
                  bulkProgress.total > 0 && (
                    <div className="mb-6 p-4 bg-green-50 border border-green-200 rounded-md">
                      <div className="flex items-center">
                        <div className="text-green-600 mr-3">
                          <svg
                            className="h-6 w-6"
                            fill="none"
                            stroke="currentColor"
                            viewBox="0 0 24 24"
                          >
                            <path
                              strokeLinecap="round"
                              strokeLinejoin="round"
                              strokeWidth={2}
                              d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z"
                            />
                          </svg>
                        </div>
                        <div>
                          <p className="text-green-800 font-medium">
                            All files processed successfully!
                          </p>
                          <p className="text-green-600 text-sm mt-1">
                            {bulkResults.length} files completed, {bulkErrors.length} errors
                          </p>
                        </div>
                      </div>
                    </div>
                  )}

                {isProcessing && (
                  <div className="mb-6 p-4 bg-blue-50 border border-blue-200 rounded-md">
                    <div className="flex items-center justify-between">
                      <div className="flex items-center">
                        <div className="animate-spin rounded-full h-4 w-4 border-b-2 border-blue-600 mr-3"></div>
                        <span className="text-blue-800">{processingMessage}</span>
                      </div>
                      {currentProcessingFile && (
                        <span className="text-sm text-blue-600 font-medium">
                          Currently: {currentProcessingFile}
                        </span>
                      )}
                    </div>
                  </div>
                )}

                {/* Progress Summary */}
                <div className="mb-6 p-4 bg-gray-50 rounded-md">
                  <div className="flex justify-between items-center">
                    <span className="text-sm font-medium text-gray-700">
                      Progress: {bulkProgress.completed} / {bulkProgress.total} files
                    </span>
                    <span className="text-sm text-gray-500">
                      {Math.round((bulkProgress.completed / bulkProgress.total) * 100)}% complete
                    </span>
                  </div>
                  {/* Progress Bar */}
                  <div className="mt-2 w-full bg-gray-200 rounded-full h-2">
                    <div
                      className={`h-2 rounded-full transition-all duration-300 ${
                        bulkProgress.completed === bulkProgress.total && bulkProgress.total > 0
                          ? 'bg-green-600'
                          : 'bg-blue-600'
                      }`}
                      style={{ width: `${(bulkProgress.completed / bulkProgress.total) * 100}%` }}
                    ></div>
                  </div>
                </div>

                {/* Real-time File Progress */}
                {fileProgress.size > 0 && (
                  <div className="mb-6">
                    <h4 className="font-medium text-gray-900 mb-3">
                      File Progress: (Total: {fileProgress.size} files)
                    </h4>
                    <div className="space-y-2">
                      {(() => {
                        const entries = Array.from(fileProgress.entries())
                        console.log('All file progress entries:', entries)
                        console.log(
                          'Entry keys types:',
                          entries.map(([k]) => typeof k)
                        )
                        // Sort by numeric value of string keys
                        const sorted = entries.sort(([a], [b]) => Number(a) - Number(b))
                        console.log('Sorted entries:', sorted)
                        return sorted.map(([index, progress]) => {
                          console.log(
                            `Rendering file index ${index} (type: ${typeof index}):`,
                            progress
                          )
                          return (
                            <div
                              key={`file-${index}`}
                              className={`p-3 rounded-md border ${
                                progress.status === 'success'
                                  ? 'bg-green-50 border-green-200'
                                  : progress.status === 'failed'
                                    ? 'bg-red-50 border-red-200'
                                    : progress.status === 'pending'
                                      ? 'bg-gray-50 border-gray-200'
                                      : 'bg-blue-50 border-blue-200'
                              }`}
                            >
                              <div className="flex items-center justify-between">
                                <div className="flex items-center">
                                  {progress.status === 'pending' && (
                                    <div className="w-3 h-3 rounded-full bg-gray-400 mr-2"></div>
                                  )}
                                  {progress.status === 'processing' && (
                                    <div className="animate-spin rounded-full h-3 w-3 border-b-2 border-blue-600 mr-2"></div>
                                  )}
                                  {progress.status === 'success' && (
                                    <div className="text-green-600 mr-2">✓</div>
                                  )}
                                  {progress.status === 'failed' && (
                                    <div className="text-red-600 mr-2">✗</div>
                                  )}
                                  <span className="text-sm font-medium text-gray-700">
                                    {progress.message}
                                  </span>
                                </div>
                                <span
                                  className={`text-xs px-2 py-1 rounded-full ${
                                    progress.status === 'success'
                                      ? 'bg-green-100 text-green-800'
                                      : progress.status === 'failed'
                                        ? 'bg-red-100 text-red-800'
                                        : progress.status === 'pending'
                                          ? 'bg-gray-100 text-gray-800'
                                          : 'bg-blue-100 text-blue-800'
                                  }`}
                                >
                                  {progress.status}
                                </span>
                              </div>
                              {progress.error && (
                                <div className="mt-1 text-xs text-red-600">
                                  Error: {progress.error}
                                </div>
                              )}
                            </div>
                          )
                        })
                      })()}
                    </div>
                  </div>
                )}

                {/* Successfully Processed Files */}
                {bulkResults.length > 0 && (
                  <div className="mb-6">
                    <div className="flex justify-between items-center mb-3">
                      <h4 className="font-medium text-gray-900">Successfully Processed Files:</h4>
                      <button
                        onClick={handleDownloadAll}
                        disabled={bulkResults.every((result) =>
                          downloadedFiles.has(result.excel_filename)
                        )}
                        className={`px-4 py-2 rounded-md text-sm font-medium transition-colors ${
                          bulkResults.every((result) => downloadedFiles.has(result.excel_filename))
                            ? 'bg-gray-100 text-gray-500 cursor-not-allowed'
                            : 'bg-green-600 text-white hover:bg-green-700'
                        }`}
                      >
                        {bulkResults.every((result) => downloadedFiles.has(result.excel_filename))
                          ? '✓ All Downloaded'
                          : `Download All (${bulkResults.length})`}
                      </button>
                    </div>
                    <div className="space-y-2">
                      {bulkResults.map((result, index) => (
                        <div
                          key={index}
                          className="flex items-center justify-between bg-white p-3 rounded-md border"
                        >
                          <div className="text-sm text-gray-700">
                            <div className="font-medium">{result.filename}</div>
                            <div className="text-gray-500">→ {result.excel_filename}</div>
                          </div>
                          <button
                            onClick={() => handleIndividualDownload(result)}
                            disabled={downloadedFiles.has(result.excel_filename)}
                            className={`px-3 py-1 rounded-md text-sm font-medium transition-colors ${
                              downloadedFiles.has(result.excel_filename)
                                ? 'bg-gray-100 text-gray-500 cursor-not-allowed'
                                : 'bg-blue-600 text-white hover:bg-blue-700'
                            }`}
                          >
                            {downloadedFiles.has(result.excel_filename)
                              ? '✓ Downloaded'
                              : 'Download'}
                          </button>
                        </div>
                      ))}
                    </div>
                  </div>
                )}

                {/* Failed Files */}
                {bulkErrors.length > 0 && (
                  <div className="mb-6">
                    <h4 className="font-medium text-red-800 mb-3">Failed Files:</h4>
                    <div className="space-y-2">
                      {bulkErrors.map((error, index) => (
                        <div key={index} className="bg-red-50 p-3 rounded-md border border-red-200">
                          <div className="text-sm text-red-700">
                            <div className="font-medium">{error.filename}</div>
                            <div className="text-red-600">{error.error}</div>
                          </div>
                        </div>
                      ))}
                    </div>
                  </div>
                )}

                {/* Summary */}
                <div className="mt-6 pt-4 border-t border-gray-200">
                  <div className="flex justify-between items-center">
                    <div className="text-sm text-gray-600">
                      <span className="text-green-600 font-medium">
                        {bulkResults.length} successful
                      </span>
                      {bulkErrors.length > 0 && (
                        <span className="text-red-600 font-medium ml-2">
                          , {bulkErrors.length} failed
                        </span>
                      )}
                      {bulkResults.length > 0 && (
                        <span className="text-blue-600 font-medium ml-2">
                          , {Array.from(downloadedFiles).length} downloaded
                        </span>
                      )}
                    </div>
                    <button
                      onClick={handleReset}
                      className="bg-blue-600 text-white px-4 py-2 rounded-md hover:bg-blue-700 transition-colors"
                    >
                      Process More Files
                    </button>
                  </div>
                </div>
              </div>
            </div>
          )
        default:
          return (
            <div className="space-y-6">
              <div className="bg-white p-6 rounded-lg shadow-sm border">
                <h2 className="text-xl font-semibold text-gray-900 mb-4">Bulk Processing</h2>
                <p className="text-gray-600">Please upload files to start processing.</p>
              </div>
            </div>
          )
      }
    }

    // Default return for any unhandled cases
    return (
      <div className="space-y-6">
        <div className="bg-white p-6 rounded-lg shadow-sm border">
          <h2 className="text-xl font-semibold text-gray-900 mb-4">Processing</h2>
          <p className="text-gray-600">Please select a processing mode.</p>
        </div>
      </div>
    )
  }

  return (
    <div className="min-h-screen bg-gray-50">
      <div className="max-w-4xl mx-auto py-8 px-4">
        <div className="text-center mb-8">
          <h1 className="text-3xl font-bold text-gray-900 mb-2">Word to Excel Processor</h1>
          <p className="text-gray-600">
            Extract market research data and generate Excel files with preserved macros
          </p>
        </div>

        {/* Tab Navigation */}
        <div className="flex space-x-1 bg-white p-1 rounded-lg shadow-sm border mb-6">
          <button
            onClick={() => {
              setActiveTab('single')
              handleReset()
            }}
            className={`flex-1 py-2 px-4 rounded-md text-sm font-medium transition-colors ${
              activeTab === 'single'
                ? 'bg-blue-600 text-white'
                : 'text-gray-600 hover:text-gray-900'
            }`}
          >
            Single File
          </button>
          <button
            onClick={() => {
              setActiveTab('bulk')
              handleReset()
            }}
            className={`flex-1 py-2 px-4 rounded-md text-sm font-medium transition-colors ${
              activeTab === 'bulk' ? 'bg-blue-600 text-white' : 'text-gray-600 hover:text-gray-900'
            }`}
          >
            Bulk Processing
          </button>
        </div>

        {/* Error Display */}
        {error && (
          <div className="mb-6 p-4 bg-red-50 border border-red-200 rounded-md">
            <div className="flex">
              <div className="text-red-600 text-xl mr-3">⚠</div>
              <div>
                <h3 className="text-red-800 font-medium">Something Went Wrong</h3>
                <p className="text-red-700 text-sm mt-1">{error}</p>
                <button
                  onClick={() => setError('')}
                  className="text-red-600 hover:text-red-800 text-sm mt-2 underline"
                >
                  Dismiss
                </button>
              </div>
            </div>
          </div>
        )}

        {/* Processing Step Display */}
        {processingStep && (
          <div className="mb-6 p-4 bg-blue-50 border border-blue-200 rounded-md">
            <div className="flex items-center">
              <div className="animate-spin rounded-full h-4 w-4 border-b-2 border-blue-600 mr-3"></div>
              <span className="text-blue-800">{processingStep}</span>
            </div>
          </div>
        )}

        {/* Main Content */}
        {renderContent()}
      </div>
    </div>
  )
}
