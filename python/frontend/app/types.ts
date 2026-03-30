export interface MarketData {
  market_name: string
  base_year: number
  start_year: number
  end_year: number
  size_base_raw: string
  size_forecast_raw: string
  cagr_percent_display: string
  currency_unit: string
  size_base_usd_mn?: number
  size_forecast_usd_mn?: number
  cagr_percent_num?: number
  driver_1?: string
  driver_2?: string
  restraint_1?: string
  restraint_2?: string
}

export interface SimpleMarketData {
  market_name: string
  market_size: string
  growth_rate: string
  key_drivers: string
  challenges: string
  opportunities: string
}

export interface SegmentationItem {
  header: string
  items: string[]
  shares?: number[]
}

export interface SegmentationData {
  cat_1: SegmentationItem
  cat_2: SegmentationItem
  cat_3: SegmentationItem
  cat_4?: SegmentationItem
  cat_5?: SegmentationItem
}

export interface KeyPlayersData {
  header: string
  players: string[]
}

export interface ExtractedData {
  market: MarketData | SimpleMarketData
  segments: SegmentationData
  players: KeyPlayersData | string[]
}

export interface ProcessingResult {
  success: boolean
  data?: ExtractedData
  error?: string
  confidence?: number
}

export interface FileUploadResponse {
  success: boolean
  fileId?: string
  error?: string
  message?: string
}
