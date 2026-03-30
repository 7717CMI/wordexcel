'use client'

import { useState, useRef } from 'react'

interface FileUploadProps {
  onFileUploaded: (fileId: string, quickProcess: boolean) => void
}

export default function FileUpload({ onFileUploaded }: FileUploadProps) {
  const [isUploading, setIsUploading] = useState(false)
  const [dragActive, setDragActive] = useState(false)
  const [error, setError] = useState<string>('')
  const [quickProcess, setQuickProcess] = useState(true) // Default to checked
  const fileInputRef = useRef<HTMLInputElement>(null)

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

    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      handleFile(e.dataTransfer.files[0])
    }
  }

  const handleFileSelect = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      handleFile(e.target.files[0])
    }
  }

  const handleFile = async (file: File) => {
    // Validate file type
    if (!file.name.toLowerCase().endsWith('.docx') && !file.name.toLowerCase().endsWith('.doc')) {
      setError('Please select a Word document (.docx or .doc)')
      return
    }

    // Validate file size (50MB limit)
    if (file.size > 50 * 1024 * 1024) {
      setError('File size must be less than 50MB')
      return
    }

    setIsUploading(true)
    setError('')

    try {
      const formData = new FormData()
      formData.append('file', file)

      const response = await fetch('/api/upload', {
        method: 'POST',
        body: formData,
      })

      if (!response.ok) {
        throw new Error('Upload failed')
      }

      const result = await response.json()

      if (result.success && result.fileId) {
        onFileUploaded(result.fileId, quickProcess)
      } else {
        throw new Error(result.error || 'Upload failed')
      }
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Upload failed')
    } finally {
      setIsUploading(false)
    }
  }

  const handleClick = () => {
    fileInputRef.current?.click()
  }

  return (
    <div className="card">
      <div className="text-center">
        <h2 className="text-2xl font-semibold text-gray-900 mb-4">Upload Word Document</h2>
        <p className="text-gray-600 mb-6">
          Upload a Word document (.docx or .doc) containing market research data
        </p>

        {/* Error Display */}
        {error && (
          <div className="bg-red-50 border border-red-200 rounded-md p-4 mb-6">
            <div className="text-sm text-red-700">{error}</div>
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
            <div className="text-lg text-gray-900 mb-2">
              {isUploading ? 'Uploading...' : 'Drop your Word document here'}
            </div>
            <p className="text-sm text-gray-500 mb-4">or click to browse files</p>
            <button
              type="button"
              onClick={handleClick}
              disabled={isUploading}
              className="btn-primary disabled:opacity-50 disabled:cursor-not-allowed"
            >
              {isUploading ? 'Uploading...' : 'Select File'}
            </button>
          </div>
        </div>

        {/* File Input */}
        <input
          ref={fileInputRef}
          type="file"
          accept=".docx,.doc"
          onChange={handleFileSelect}
          className="hidden"
        />

        {/* Quick Process Option */}
        <div className="mt-6 p-4 bg-gray-50 rounded-lg">
          <div className="flex items-start">
            <div className="flex items-center h-5">
              <input
                id="quick-process"
                type="checkbox"
                checked={quickProcess}
                onChange={(e) => setQuickProcess(e.target.checked)}
                className="focus:ring-primary-500 h-4 w-4 text-primary-600 border-gray-300 rounded"
              />
            </div>
            <div className="ml-3 text-sm">
              <label htmlFor="quick-process" className="font-medium text-gray-900 cursor-pointer">
                Quick Process (Auto Download)
              </label>
              <div className="text-gray-600 mt-1">
                {quickProcess ? (
                  <>
                    ✓ Extract data automatically
                    <br />
                    ✓ Generate Excel automatically
                    <br />✓ Download immediately
                  </>
                ) : (
                  <>
                    → Extract data automatically
                    <br />
                    → Show for review/editing
                    <br />→ Manual Excel generation
                  </>
                )}
              </div>
            </div>
          </div>
        </div>

        {/* Instructions */}
        <div className="mt-6 text-sm text-gray-500">
          <p className="mb-2">Supported formats: .docx, .doc</p>
          <p>Maximum file size: 50MB</p>
        </div>
      </div>
    </div>
  )
}
