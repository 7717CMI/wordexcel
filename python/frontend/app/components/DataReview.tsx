'use client'

import { useState } from 'react'
import { ExtractedData, MarketData, SimpleMarketData, KeyPlayersData } from '../types'

interface DataReviewProps {
  data: ExtractedData
  onDataConfirmed: (data: ExtractedData) => void
  onBack: () => void
}

export default function DataReview({ data, onDataConfirmed, onBack }: DataReviewProps) {
  const [editedData, setEditedData] = useState<ExtractedData>(data)

  // Type guards
  const isMarketData = (market: MarketData | SimpleMarketData): market is MarketData => {
    return 'base_year' in market
  }

  const isKeyPlayersData = (players: KeyPlayersData | string[]): players is KeyPlayersData => {
    return typeof players === 'object' && 'header' in players && 'players' in players
  }

  const handleMarketChange = (field: string, value: string | number) => {
    setEditedData((prev) => ({
      ...prev,
      market: {
        ...prev.market,
        [field]: value,
      },
    }))
  }

  const handleSegmentChange = (
    category: keyof ExtractedData['segments'],
    field: 'header' | 'items',
    value: string | string[]
  ) => {
    setEditedData((prev) => ({
      ...prev,
      segments: {
        ...prev.segments,
        [category]: {
          ...prev.segments[category],
          [field]: value,
        },
      },
    }))
  }

  const handlePlayersChange = (field: 'header' | 'players', value: string | string[]) => {
    if (isKeyPlayersData(editedData.players)) {
      setEditedData((prev) => ({
        ...prev,
        players: {
          ...prev.players,
          [field]: value,
        },
      }))
    }
  }

  return (
    <div className="space-y-6">
      <div className="bg-white p-6 rounded-lg shadow-sm border">
        <h2 className="text-xl font-semibold text-gray-900 mb-4">Review Extracted Data</h2>

        {/* Market Data */}
        <div className="mb-6">
          <h3 className="text-lg font-medium text-gray-900 mb-3">Market Information</h3>
          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-1">Market Name</label>
              <input
                type="text"
                value={editedData.market.market_name}
                onChange={(e) => handleMarketChange('market_name', e.target.value)}
                className="input-field"
              />
            </div>

            {/* Render MarketData fields if available */}
            {isMarketData(editedData.market) && (
              <>
                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-1">Base Year</label>
                  <input
                    type="number"
                    value={editedData.market.base_year}
                    onChange={(e) => handleMarketChange('base_year', parseInt(e.target.value))}
                    className="input-field"
                  />
                </div>
                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-1">Start Year</label>
                  <input
                    type="number"
                    value={editedData.market.start_year}
                    onChange={(e) => handleMarketChange('start_year', parseInt(e.target.value))}
                    className="input-field"
                  />
                </div>
                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-1">End Year</label>
                  <input
                    type="number"
                    value={editedData.market.end_year}
                    onChange={(e) => handleMarketChange('end_year', parseInt(e.target.value))}
                    className="input-field"
                  />
                </div>
                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-1">
                    Market Size (Base Year)
                  </label>
                  <input
                    type="text"
                    value={editedData.market.size_base_raw}
                    onChange={(e) => handleMarketChange('size_base_raw', e.target.value)}
                    className="input-field"
                    placeholder="e.g., USD 150 Mn"
                  />
                </div>
                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-1">
                    Market Size (Forecast)
                  </label>
                  <input
                    type="text"
                    value={editedData.market.size_forecast_raw}
                    onChange={(e) => handleMarketChange('size_forecast_raw', e.target.value)}
                    className="input-field"
                    placeholder="e.g., USD 290 Mn"
                  />
                </div>
                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-1">CAGR</label>
                  <input
                    type="text"
                    value={editedData.market.cagr_percent_display}
                    onChange={(e) => handleMarketChange('cagr_percent_display', e.target.value)}
                    className="input-field"
                    placeholder="e.g., 9.50%"
                  />
                </div>
                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-1">
                    Currency Unit
                  </label>
                  <input
                    type="text"
                    value={editedData.market.currency_unit}
                    onChange={(e) => handleMarketChange('currency_unit', e.target.value)}
                    className="input-field"
                    placeholder="e.g., USD"
                  />
                </div>
                <div className="md:col-span-2">
                  <label className="block text-sm font-medium text-gray-700 mb-1">
                    Driver 1 - Outline (C15)
                  </label>
                  <input
                    type="text"
                    value={editedData.market.driver_1 || ''}
                    onChange={(e) => handleMarketChange('driver_1', e.target.value)}
                    className="input-field"
                    placeholder="Short outline (5-10 words)"
                  />
                </div>
                <div className="md:col-span-2">
                  <label className="block text-sm font-medium text-gray-700 mb-1">
                    Driver 2 - Outline (C16)
                  </label>
                  <input
                    type="text"
                    value={editedData.market.driver_2 || ''}
                    onChange={(e) => handleMarketChange('driver_2', e.target.value)}
                    className="input-field"
                    placeholder="Short outline (5-10 words)"
                  />
                </div>
                <div className="md:col-span-2">
                  <label className="block text-sm font-medium text-gray-700 mb-1">
                    Restraint 1 (C17)
                  </label>
                  <textarea
                    value={editedData.market.restraint_1 || ''}
                    onChange={(e) => handleMarketChange('restraint_1', e.target.value)}
                    className="input-field"
                    rows={2}
                    placeholder="First key market restraint/challenge..."
                  />
                </div>
                <div className="md:col-span-2">
                  <label className="block text-sm font-medium text-gray-700 mb-1">
                    Restraint 2 (C18)
                  </label>
                  <textarea
                    value={editedData.market.restraint_2 || ''}
                    onChange={(e) => handleMarketChange('restraint_2', e.target.value)}
                    className="input-field"
                    rows={2}
                    placeholder="Second key market restraint/challenge..."
                  />
                </div>
              </>
            )}

            {/* Render SimpleMarketData fields if available */}
            {!isMarketData(editedData.market) && (
              <>
                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-1">
                    Market Size
                  </label>
                  <input
                    type="text"
                    value={editedData.market.market_size}
                    onChange={(e) => handleMarketChange('market_size', e.target.value)}
                    className="input-field"
                  />
                </div>
                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-1">
                    Growth Rate
                  </label>
                  <input
                    type="text"
                    value={editedData.market.growth_rate}
                    onChange={(e) => handleMarketChange('growth_rate', e.target.value)}
                    className="input-field"
                  />
                </div>
                <div className="md:col-span-2">
                  <label className="block text-sm font-medium text-gray-700 mb-1">
                    Key Drivers
                  </label>
                  <textarea
                    value={editedData.market.key_drivers}
                    onChange={(e) => handleMarketChange('key_drivers', e.target.value)}
                    className="input-field"
                    rows={3}
                  />
                </div>
                <div className="md:col-span-2">
                  <label className="block text-sm font-medium text-gray-700 mb-1">Challenges</label>
                  <textarea
                    value={editedData.market.challenges}
                    onChange={(e) => handleMarketChange('challenges', e.target.value)}
                    className="input-field"
                    rows={3}
                  />
                </div>
                <div className="md:col-span-2">
                  <label className="block text-sm font-medium text-gray-700 mb-1">
                    Opportunities
                  </label>
                  <textarea
                    value={editedData.market.opportunities}
                    onChange={(e) => handleMarketChange('opportunities', e.target.value)}
                    className="input-field"
                    rows={3}
                  />
                </div>
              </>
            )}
          </div>
        </div>

        {/* Segmentation Data */}
        <div className="mb-6">
          <h3 className="text-lg font-medium text-gray-900 mb-3">Market Segmentation</h3>
          <div className="space-y-4">
            {Object.entries(editedData.segments).map(([key, segment]) => (
              <div key={key} className="border rounded-lg p-4">
                <div className="mb-2">
                  <label className="block text-sm font-medium text-gray-700 mb-1">
                    {key.replace('_', ' ').toUpperCase()} Header
                  </label>
                  <input
                    type="text"
                    value={segment.header}
                    onChange={(e) =>
                      handleSegmentChange(
                        key as keyof ExtractedData['segments'],
                        'header',
                        e.target.value
                      )
                    }
                    className="input-field"
                  />
                </div>
                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-1">
                    Items (one per line)
                  </label>
                  <textarea
                    value={segment.items.join('\n')}
                    onChange={(e) =>
                      handleSegmentChange(
                        key as keyof ExtractedData['segments'],
                        'items',
                        e.target.value.split('\n')
                      )
                    }
                    className="input-field"
                    rows={4}
                  />
                </div>
              </div>
            ))}
          </div>
        </div>

        {/* Key Players */}
        <div className="mb-6">
          <h3 className="text-lg font-medium text-gray-900 mb-3">Key Players</h3>
          {isKeyPlayersData(editedData.players) ? (
            <div className="space-y-4">
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-1">Header</label>
                <input
                  type="text"
                  value={editedData.players.header}
                  onChange={(e) => handlePlayersChange('header', e.target.value)}
                  className="input-field"
                />
              </div>
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-1">
                  Players (one per line)
                </label>
                <textarea
                  value={editedData.players.players.join('\n')}
                  onChange={(e) => handlePlayersChange('players', e.target.value.split('\n'))}
                  className="input-field"
                  rows={4}
                />
              </div>
            </div>
          ) : (
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-1">
                Players (one per line)
              </label>
              <textarea
                value={Array.isArray(editedData.players) ? editedData.players.join('\n') : ''}
                onChange={(e) => {
                  const players = e.target.value.split('\n')
                  setEditedData((prev) => ({
                    ...prev,
                    players: players,
                  }))
                }}
                className="input-field"
                rows={4}
              />
            </div>
          )}
        </div>

        {/* Action Buttons */}
        <div className="flex justify-between">
          <button onClick={onBack} className="btn-secondary">
            Back
          </button>
          <button onClick={() => onDataConfirmed(editedData)} className="btn-primary">
            Confirm & Generate Excel
          </button>
        </div>
      </div>
    </div>
  )
}
