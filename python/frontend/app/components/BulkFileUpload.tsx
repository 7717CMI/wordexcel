'use client'

import { useState, useRef } from 'react'

interface BulkFileUploadProps {
  onBulkUploaded: (files: File[]) => void
  isProcessing?: boolean
  fileProgress?: Map<string, { status: string; message: string; error?: string }>
  onFilesSelected?: (files: File[]) => void
}

export default function BulkFileUpload({
  onBulkUploaded,
  isProcessing: externalProcessing,
  fileProgress,
  onFilesSelected,
}: BulkFileUploadProps) {
  const [isUploading] = useState(false)
  const [isProcessing, setIsProcessing] = useState(false)
  const [dragActive, setDragActive] = useState(false)
  const [error, setError] = useState<string>('')
  const [selectedFiles, setSelectedFiles] = useState<File[]>([])
  const fileInputRef = useRef<HTMLInputElement>(null)

  // Use external processing state if provided, otherwise use internal state
  const processingState = externalProcessing !== undefined ? externalProcessing : isProcessing

  const handleDrag = (e: React.DragEvent) => {
    e.preventDefault()
    e.stopPropagation()
    if (e.type === 'dragenter' || e.type === 'dragover') {
      setDragActive(true)
    } else if (e.type === 'dragleave') {
      setDragActive(false)
    }
  }

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault()
    e.stopPropagation()
    setDragActive(false)

    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      handleFiles(Array.from(e.dataTransfer.files))
    }
  }

  const handleFileSelect = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files.length > 0) {
      handleFiles(Array.from(e.target.files))
    }
  }

  const handleFiles = (files: File[]) => {
    // Validate total file count first
    if (files.length > 50) {
      setError(
        `Too many files selected. Maximum 50 files allowed, but ${files.length} files were selected.`
      )
      return
    }

    // Validate files
    const validFiles: File[] = []
    const errors: string[] = []

    for (const file of files) {
      // Validate file type
      if (!file.name.toLowerCase().endsWith('.docx') && !file.name.toLowerCase().endsWith('.doc')) {
        errors.push(`${file.name}: Only Word documents (.docx or .doc) are allowed`)
        continue
      }

      // Validate file size (50MB limit)
      if (file.size > 50 * 1024 * 1024) {
        errors.push(`${file.name}: File size must be less than 50MB`)
        continue
      }

      validFiles.push(file)
    }

    if (errors.length > 0) {
      setError(errors.join('\n'))
    } else {
      setError('')
    }

    if (validFiles.length > 0) {
      setSelectedFiles(validFiles)
      // Notify parent component that new files are selected (to clear progress)
      if (onFilesSelected) {
        onFilesSelected(validFiles)
      }
    }
  }

  const handleClick = () => {
    fileInputRef.current?.click()
  }

  const handleStartProcessing = () => {
    if (selectedFiles.length === 0) {
      setError('Please select at least one file')
      return
    }

    if (selectedFiles.length > 50) {
      setError('Maximum 50 files allowed per batch')
      return
    }

    // Check if all files have been processed (Process Another Files scenario)
    if (fileProgress && selectedFiles.length > 0) {
      const completedCount = Array.from(fileProgress.values()).filter(
        (progress) => progress.status === 'success' || progress.status === 'failed'
      ).length

      if (completedCount === selectedFiles.length) {
        // Reset for new batch
        setSelectedFiles([])
        setError('')
        // Clear the file input
        if (fileInputRef.current) {
          fileInputRef.current.value = ''
        }
        return
      }
    }

    setIsProcessing(true)
    setError('')
    onBulkUploaded(selectedFiles)
  }

  const removeFile = (index: number) => {
    const newFiles = selectedFiles.filter((_, i) => i !== index)
    setSelectedFiles(newFiles)
  }

  return (
    <div className="card">
      <div className="text-center">
        <h2 className="text-2xl font-semibold text-gray-900 mb-4">Bulk Upload Word Documents</h2>
        <p className="text-gray-600 mb-6">
          Upload multiple Word documents (.docx or .doc) for parallel processing (max 50 files)
        </p>

        {/* Error Display */}
        {error && (
          <div className="bg-red-50 border border-red-200 rounded-md p-4 mb-6">
            <div className="text-sm text-red-700 whitespace-pre-line">{error}</div>
          </div>
        )}

        {/* Upload Area */}
        <div
          className={`border-2 border-dashed rounded-lg p-8 transition-colors duration-200 ${
            dragActive
              ? 'border-primary-500 bg-primary-50'
              : 'border-gray-300 hover:border-primary-400'
          }`}
          onDragEnter={handleDrag}
          onDragLeave={handleDrag}
          onDragOver={handleDrag}
          onDrop={handleDrop}
        >
          <div className="text-center">
            <svg
              className="mx-auto h-12 w-12 text-gray-400 mb-4"
              stroke="currentColor"
              fill="none"
              viewBox="0 0 48 48"
            >
              <path
                d="M28 8H12a4 4 0 00-4 4v20m32-12v8m0 0v8a4 4 0 01-4 4H12a4 4 0 01-4-4v-4m32-4l-3.172-3.172a4 4 0 00-5.656 0L28 28M8 32l9.172-9.172a4 4 0 015.656 0L28 28m0 0l4 4m4-24h8m-4-4v8m-12 4h.02"
                strokeWidth={2}
                strokeLinecap="round"
                strokeLinejoin="round"
              />
            </svg>
            <div className="text-lg text-gray-900 mb-2">Drop multiple Word documents here</div>
            <p className="text-sm text-gray-500 mb-4">or click to browse files</p>
            <button
              type="button"
              onClick={handleClick}
              disabled={isUploading}
              className="btn-primary disabled:opacity-50 disabled:cursor-not-allowed"
            >
              {isUploading ? 'Uploading...' : 'Select Files'}
            </button>
          </div>
        </div>

        {/* File Input */}
        <input
          ref={fileInputRef}
          type="file"
          accept=".docx,.doc"
          multiple
          onChange={handleFileSelect}
          className="hidden"
        />

        {/* Selected Files */}
        {selectedFiles.length > 0 && (
          <div className="mt-6">
            <h3 className="text-lg font-medium text-gray-900 mb-3">
              Selected Files ({selectedFiles.length})
            </h3>

            {/* Processing Status */}
            {processingState && (
              <div className="mb-4 p-3 bg-blue-50 border border-blue-200 rounded-md">
                <div className="flex items-center">
                  <div className="animate-spin rounded-full h-4 w-4 border-b-2 border-blue-600 mr-3"></div>
                  <span className="text-blue-800 text-sm font-medium">
                    Files are being processed in parallel...
                  </span>
                </div>
              </div>
            )}

            <div className="space-y-2 max-h-40 overflow-y-auto">
              {selectedFiles.map((file, index) => {
                const progress = fileProgress?.get(String(index))
                const getStatusIcon = () => {
                  if (!progress) {
                    return (
                      <svg
                        className="h-5 w-5 text-blue-500 mr-2"
                        fill="none"
                        stroke="currentColor"
                        viewBox="0 0 24 24"
                      >
                        <path
                          strokeLinecap="round"
                          strokeLinejoin="round"
                          strokeWidth={2}
                          d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"
                        />
                      </svg>
                    )
                  }

                  switch (progress.status) {
                    case 'processing':
                      return (
                        <div className="animate-spin rounded-full h-5 w-5 border-b-2 border-blue-600 mr-2"></div>
                      )
                    case 'success':
                      return (
                        <div className="text-green-600 mr-2">
                          <svg
                            className="h-5 w-5"
                            fill="none"
                            stroke="currentColor"
                            viewBox="0 0 24 24"
                          >
                            <path
                              strokeLinecap="round"
                              strokeLinejoin="round"
                              strokeWidth={2}
                              d="M5 13l4 4L19 7"
                            />
                          </svg>
                        </div>
                      )
                    case 'failed':
                      return (
                        <div className="text-red-600 mr-2">
                          <svg
                            className="h-5 w-5"
                            fill="none"
                            stroke="currentColor"
                            viewBox="0 0 24 24"
                          >
                            <path
                              strokeLinecap="round"
                              strokeLinejoin="round"
                              strokeWidth={2}
                              d="M6 18L18 6M6 6l12 12"
                            />
                          </svg>
                        </div>
                      )
                    default:
                      return (
                        <svg
                          className="h-5 w-5 text-blue-500 mr-2"
                          fill="none"
                          stroke="currentColor"
                          viewBox="0 0 24 24"
                        >
                          <path
                            strokeLinecap="round"
                            strokeLinejoin="round"
                            strokeWidth={2}
                            d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"
                          />
                        </svg>
                      )
                  }
                }

                const getStatusText = () => {
                  if (!progress) return ''

                  switch (progress.status) {
                    case 'processing':
                      return 'Processing...'
                    case 'success':
                      return 'Completed ✓'
                    case 'failed':
                      return 'Failed ✗'
                    default:
                      return ''
                  }
                }

                const getStatusColor = () => {
                  if (!progress) return 'text-gray-500'

                  switch (progress.status) {
                    case 'processing':
                      return 'text-blue-600'
                    case 'success':
                      return 'text-green-600'
                    case 'failed':
                      return 'text-red-600'
                    default:
                      return 'text-gray-500'
                  }
                }

                return (
                  <div
                    key={index}
                    className={`flex items-center justify-between p-3 rounded-md ${
                      progress?.status === 'success'
                        ? 'bg-green-50 border border-green-200'
                        : progress?.status === 'failed'
                          ? 'bg-red-50 border border-red-200'
                          : progress?.status === 'processing'
                            ? 'bg-blue-50 border border-blue-200'
                            : 'bg-gray-50'
                    }`}
                  >
                    <div className="flex items-center flex-1">
                      {getStatusIcon()}
                      <div className="flex-1">
                        <span className="text-sm text-gray-900">{file.name}</span>
                        <span className="text-xs text-gray-500 ml-2">
                          ({(file.size / 1024 / 1024).toFixed(1)} MB)
                        </span>
                        {progress && (
                          <div className={`text-xs mt-1 ${getStatusColor()}`}>
                            {getStatusText()}
                          </div>
                        )}
                      </div>
                    </div>
                    <button
                      onClick={() => removeFile(index)}
                      className={`${
                        progress?.status === 'processing'
                          ? 'text-gray-400 cursor-not-allowed'
                          : 'text-red-500 hover:text-red-700'
                      }`}
                      disabled={progress?.status === 'processing'}
                    >
                      <svg
                        className="h-4 w-4"
                        fill="none"
                        stroke="currentColor"
                        viewBox="0 0 24 24"
                      >
                        <path
                          strokeLinecap="round"
                          strokeLinejoin="round"
                          strokeWidth={2}
                          d="M6 18L18 6M6 6l12 12"
                        />
                      </svg>
                    </button>
                  </div>
                )
              })}
            </div>

            <button
              onClick={handleStartProcessing}
              disabled={(() => {
                if (isUploading || selectedFiles.length === 0) return true
                if (processingState) return true

                // Allow clicking if all files are completed
                if (fileProgress && selectedFiles.length > 0) {
                  const completedCount = Array.from(fileProgress.values()).filter(
                    (progress) => progress.status === 'success' || progress.status === 'failed'
                  ).length

                  if (completedCount === selectedFiles.length) {
                    return false // Allow clicking "Process Another Files"
                  }
                }

                return false
              })()}
              className="mt-4 w-full btn-primary disabled:opacity-50 disabled:cursor-not-allowed"
            >
              {(() => {
                if (processingState) {
                  return '🔄 Processing Files...'
                }

                if (isUploading) {
                  return 'Processing...'
                }

                // Check if all files have been processed
                if (fileProgress && selectedFiles.length > 0) {
                  const completedCount = Array.from(fileProgress.values()).filter(
                    (progress) => progress.status === 'success' || progress.status === 'failed'
                  ).length

                  if (completedCount === selectedFiles.length) {
                    return 'Process Another Files'
                  }
                }

                return `Start Processing ${selectedFiles.length} Files`
              })()}
            </button>
          </div>
        )}

        {/* Instructions */}
        <div className="mt-6 text-sm text-gray-500">
          <p className="mb-2">Supported formats: .docx, .doc</p>
          <p className="mb-2">Maximum file size: 50MB per file</p>
          <p>Maximum files per batch: 50</p>
        </div>
      </div>
    </div>
  )
}
