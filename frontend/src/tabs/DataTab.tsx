import { useState, useCallback, useRef, useEffect } from 'react'
import { Save, FileImage, Type, Mail, CheckCircle2, FileSpreadsheet, Move, Lock, LogOut } from 'lucide-react'
import {
  BASE,
  uploadExcel, uploadTemplate, clearExcel,
  getColumns, saveColumns,
  getActiveType, setActiveType,
  getCoordinatesPreview, getTextFields, saveTextFields,
  getEmailAuthStatus, setEmailAuth, signOutEmailAuth,
  getDriveAuthStatus, uploadDriveCredentials, deleteDriveCredentials, saveDriveRootFolder,
  uploadOAuthClientSecrets, getOAuthUrl, getOAuthFlowStatus, revokeOAuthToken,
  type ExcelUploadResult, type ColumnMapping, type PocGroup,
} from '../api'

/** Derive the raw backend base URL (without trailing slash) for OAuth redirect URIs. */
const getBackendUrl = (): string => {
  const apiBase = import.meta.env.VITE_API_BASE_URL
  if (apiBase) return apiBase.replace(/\/$/, '')
  return 'http://localhost:8000'
}


const CERT_TYPES = [
  { id: 'participation', label: 'Participation Certificate' },
  { id: 'winner', label: 'Winners Certificate' },
  { id: 'school', label: 'School Certificate' },
  { id: 'volunteer', label: 'Volunteer Certificate' }
] as const;
type CertType = typeof CERT_TYPES[number]['id'];

const getKeysForType = (type: CertType) => {
  if (type === 'volunteer') return ['volunteer_name', 'volunteer_email'];
  if (type === 'school') return ['school', 'poc_email', 'poc_name'];
  if (type === 'winner') return ['student_name', 'event_name', 'school', 'poc_email', 'poc_name', 'position'];
  return ['student_name', 'event_name', 'school', 'poc_email', 'poc_name'];
}

// Fields the user maps from Excel columns to our internal names
const INTERNAL_KEYS = ['student_name', 'school', 'event_name', 'poc_email', 'poc_name', 'position', 'volunteer_name', 'volunteer_email'] as const
type InternalKey = typeof INTERNAL_KEYS[number]

const KEY_LABELS: Record<InternalKey, string> = {
  student_name: 'Student Name ★',
  school: 'School / College',
  event_name: 'Event Name',
  poc_email: 'POC Email ★',
  poc_name: 'POC Name',
  position: 'Podium Position (1, 2, 3) ★',
  volunteer_name: 'Volunteer Name ★',
  volunteer_email: 'Volunteer Email ★',
}

const KEY_HINTS: Record<InternalKey, string> = {
  student_name: 'Required — printed on certificate',
  school: 'Optional — printed on certificate',
  event_name: 'Optional — printed on certificate',
  poc_email: 'Required — used to group students into clusters',
  poc_name: 'Optional — used for email greeting name',
  position: 'Required for Winners — values must be 1, 2, or 3',
  volunteer_name: 'Required — printed on volunteer certificate',
  volunteer_email: 'Required — recipient email address',
}

// ── Font Library ──────────────────────────────────────────────────────────────
const FONT_LIBRARY = [
  { label: 'DejaVu Sans',         family: 'DejaVu Sans, sans-serif',          serverPath: 'sample_data/DejaVuSans.ttf' },
  { label: 'DejaVu Sans Bold',    family: 'DejaVu Sans, sans-serif',          serverPath: 'sample_data/DejaVuSans-Bold.ttf' },
  { label: 'Montserrat SemiBold', family: "'Montserrat', sans-serif",         serverPath: 'sample_data/Montserrat-SemiBold.ttf' },
  { label: 'Dancing Script',      family: "'Dancing Script', cursive",        serverPath: 'sample_data/DancingScript-Regular.ttf' },
  { label: 'Playfair Display',    family: "'Playfair Display', serif",        serverPath: 'sample_data/PlayfairDisplay-SemiBold.ttf' },
  { label: 'Cinzel',              family: "'Cinzel', serif",            serverPath: 'sample_data/DejaVuSans.ttf' },
  { label: 'EB Garamond',         family: "'EB Garamond', serif",       serverPath: 'sample_data/DejaVuSans.ttf' },
  { label: 'Libre Baskerville',   family: "'Libre Baskerville', serif", serverPath: 'sample_data/DejaVuSans.ttf' },
  { label: 'Merriweather',        family: "'Merriweather', serif",      serverPath: 'sample_data/DejaVuSans.ttf' },
  { label: 'Raleway',             family: "'Raleway', sans-serif",      serverPath: 'sample_data/DejaVuSans.ttf' },
  { label: 'Montserrat',          family: "'Montserrat', sans-serif",   serverPath: 'sample_data/DejaVuSans.ttf' },
  { label: 'Open Sans',           family: "'Open Sans', sans-serif",    serverPath: 'sample_data/DejaVuSans.ttf' },
  { label: 'Oswald',              family: "'Oswald', sans-serif",       serverPath: 'sample_data/DejaVuSans.ttf' },
  { label: 'Roboto',              family: "'Roboto', sans-serif",       serverPath: 'sample_data/DejaVuSans.ttf' },
  { label: 'Inter',               family: "'Inter', sans-serif",        serverPath: 'sample_data/DejaVuSans.ttf' },
  { label: 'Lato',                family: "'Lato', sans-serif",         serverPath: 'sample_data/DejaVuSans.ttf' },
]

// ── Snap constants ─────────────────────────────────────────────────────────────
/** Distance in canvas-pixels within which a snap axis activates/holds */
const SNAP_THRESHOLD = 5
/** Guide line color (magenta as specified) */
const SNAP_COLOR = '#FF00EA'

// ── Types ──────────────────────────────────────────────────────────────────────

/** One candidate axis along which a field can snap */
interface SnapAxis {
  /** Canvas-coordinate value (pixels in the original image coordinate space) */
  canvasValue: number
  /** Percentage position 0–100 for CSS guide line placement */
  pct: number
}

/** Pre-computed snap targets stored once per drag gesture */
interface SnapTargets {
  /** Vertical snap axes → produce vertical guide lines (guide on X axis) */
  x: SnapAxis[]
  /** Horizontal snap axes → produce horizontal guide lines (guide on Y axis) */
  y: SnapAxis[]
}

/** Active guide positions in percentage coordinates */
interface GuideLines {
  /** Array of horizontal guide Y% positions (can have multiple) */
  h: number[]
  /** Array of vertical guide X% positions (can have multiple) */
  v: number[]
}

export default function DataTab() {

  // ── Cert Type state ────────────────────────────────────────────────────
  const [activeCertType, setActiveCertType] = useState<CertType>('participation')

  // ── Excel state ────────────────────────────────────────────────────────
  const [excelResult, setExcelResult] = useState<ExcelUploadResult | null>(null)
  const [excelLoading, setExcelLoading] = useState(false)
  const [excelError, setExcelError] = useState('')

  // ── Column mapping state ───────────────────────────────────────────────
  const [colMap, setColMap] = useState<ColumnMapping>({
    student_name: '', school: '', event_name: '', poc_email: '',
  })
  const [colSaving, setColSaving] = useState(false)
  const [colSaved, setColSaved] = useState(false)
  const [colGroups, setColGroups] = useState<PocGroup[]>([])

  // ── Template state ─────────────────────────────────────────────────────
  const [templateMode, setTemplateMode] = useState<'docx' | 'image'>('docx')
  const [templateLoading, setTemplateLoading] = useState(false)
  const [templatePath, setTemplatePath] = useState('')
  const [templatePath1st, setTemplatePath1st] = useState('')
  const [templatePath2nd, setTemplatePath2nd] = useState('')
  const [templatePath3rd, setTemplatePath3rd] = useState('')
  const [, setTemplateError] = useState('')
  const [previewError, setPreviewError] = useState('')

  // ── Coordinates / image fields ─────────────────────────────────────────
  const [gridImage, setGridImage] = useState('')
  const [gridLoading, setGridLoading] = useState(false)
  const [textFields, setTextFields] = useState<Record<string, any>>({})
  const [fieldsSaved, setFieldsSaved] = useState(false)
  const [gridWidth, setGridWidth] = useState(1200)
  const [gridHeight, setGridHeight] = useState(850)
  const [activeField, setActiveField] = useState<string | null>(null)
  const [draggingField, setDraggingField] = useState<string | null>(null)
  const [guides, setGuides] = useState<GuideLines>({ h: [], v: [] })
  const [isDrawingLine, setIsDrawingLine] = useState(false)
  // Live preview of the baseline being drawn: percentages in CSS space
  const [drawingPreview, setDrawingPreview] = useState<{ x1Pct: number, x2Pct: number, yPct: number } | null>(null)

  // ── Refs ───────────────────────────────────────────────────────────────
  const imgRef = useRef<HTMLImageElement>(null)
  const canvasContainerRef = useRef<HTMLDivElement>(null)

  /**
   * Ref mirror of `draggingField` state.
   * Mouse-event handlers read from this ref so they are NEVER stale closures
   * — state updates are async and would cause handlers to see the old value.
   */
  const draggingFieldRef = useRef<string | null>(null)

  /**
   * Static snap targets computed ONCE at drag-start (onPointerDown).
   * Stored in a ref so handleMouseMove reads stable data without
   * causing re-renders or stale closures.
   */
  const snapTargetsRef = useRef<SnapTargets>({ x: [], y: [] })

  /**
   * Stable ref to the latest textFields so onPointerDown closure
   * always sees the current fields without being in the dependency array.
   */
  const textFieldsRef = useRef<Record<string, any>>({})
  const drawStartRef = useRef<{ x: number, y: number } | null>(null)
  const isDrawingLineRef = useRef<boolean>(false)
  const activeFieldRef = useRef<string | null>(null)
  const [activeWinnerPos, setActiveWinnerPos] = useState<'1st' | '2nd' | '3rd'>('1st')
  const activeWinnerPosRef = useRef<'1st' | '2nd' | '3rd'>('1st')
  useEffect(() => { activeWinnerPosRef.current = activeWinnerPos }, [activeWinnerPos])
  useEffect(() => { textFieldsRef.current = textFields }, [textFields])
  useEffect(() => { activeFieldRef.current = activeField }, [activeField])

  /** Stable ref to grid dimensions for the same reason */
  const gridDimsRef = useRef({ w: 1200, h: 850 })
  useEffect(() => { gridDimsRef.current = { w: gridWidth, h: gridHeight } }, [gridWidth, gridHeight])

  // ── Auth state ─────────────────────────────────────────────────────────
  const [senderEmail, setSenderEmail] = useState('')
  const [appPassword, setAppPassword] = useState('')
  const [authSaving, setAuthSaving] = useState(false)
  const [authSaved, setAuthSaved] = useState(false)
  const [authError, setAuthError] = useState('')
  const [savedEmail, setSavedEmail] = useState<string | null>(null)
  const [signingOut, setSigningOut] = useState(false)

  // ── Drive Auth state ────────────────────────────────────────────────────
  const [driveAvailable, setDriveAvailable] = useState(false)
  const [driveCredsExist, setDriveCredsExist] = useState(false)
  const [driveUploading, setDriveUploading] = useState(false)
  const [driveError, setDriveError] = useState('')
  const [showDriveGuide, setShowDriveGuide] = useState(false)
  const [serviceAccountEmail, setServiceAccountEmail] = useState<string | null>(null)
  const [driveFolderInput, setDriveFolderInput] = useState('')
  const [driveFolderSaved, setDriveFolderSaved] = useState(false)
  const [saveFolderLoading, setSaveFolderLoading] = useState(false)
  // OAuth2
  const [oauthClientAvailable, setOauthClientAvailable] = useState(false)
  const [oauthDriveAvailable, setOauthDriveAvailable] = useState(false)
  const [oauthUserEmail, setOauthUserEmail] = useState<string | null>(null)
  const [oauthFlowStatus, setOauthFlowStatus] = useState<'idle' | 'in_progress' | 'done' | 'error'>('idle')
  const [oauthFlowError, setOauthFlowError] = useState('')
  const [oauthUploading, setOauthUploading] = useState(false)


  // ── Effects ────────────────────────────────────────────────────────────

  useEffect(() => {
    getEmailAuthStatus().then(s => {
      if (s.password_set && s.sender_email) {
        setSavedEmail(s.sender_email)
        setAuthSaved(true)
        setSenderEmail(s.sender_email)
      }
    }).catch(() => {/* ignore */})

    getDriveAuthStatus().then(d => {
      setDriveAvailable(d.drive_available)
      setDriveCredsExist(d.has_credentials)
      setServiceAccountEmail(d.service_account_email)
      if (d.root_folder_id) setDriveFolderInput(d.root_folder_id)
      setOauthClientAvailable(d.oauth_client_available)
      setOauthDriveAvailable(d.oauth_drive_available)
      setOauthUserEmail(d.oauth_user_email)
      if (d.oauth_drive_available) setOauthFlowStatus('done')
    }).catch(() => {})



    getActiveType().then(res => {
      if (res.active_type) setActiveCertType(res.active_type as CertType)
    }).catch(() => {})

    getColumns().then(res => {
      if (res.columns) {
        setColMap(res.columns)
      }
      // Only restore excel state if the backend has an actual file with detected columns
      if (res.excel_path && res.excel_columns && res.excel_columns.length > 0) {
        setExcelResult({
          excel_columns: res.excel_columns,
          groups: res.groups || [],
          excel_path: res.excel_path
        })
        if (res.groups) {
          setColGroups(res.groups)
        }
      }
    }).catch(e => console.error('Failed to load columns config:', e))
  }, [])

  // ── Hydrate template mode + path from backend config on mount ────────────
  // This ensures the Visual Editor section is visible even after a page reload,
  // and auto-loads the preview if the config already has an image template.
  useEffect(() => {
    // Fetch template mode from a lightweight config read
    fetch(`${BASE}/config/template`)
      .then(r => r.ok ? r.json() : null)
      .then(data => {
        if (!data) return
        const mode: 'docx' | 'image' = data.template_mode === 'image' ? 'image' : 'docx'
        setTemplateMode(mode)
        if (data.image_template_path_1st) setTemplatePath1st(data.image_template_path_1st)
        if (data.image_template_path_2nd) setTemplatePath2nd(data.image_template_path_2nd)
        if (data.image_template_path_3rd) setTemplatePath3rd(data.image_template_path_3rd)
        if (data.image_template_path) setTemplatePath(data.image_template_path)
        else if (data.docx_template_path) setTemplatePath(data.docx_template_path)
        // Auto-load preview if backend already has an image template configured
        if (mode === 'image' && data.image_template_path) {
          handleGridPreview()
        }
      })
      .catch(() => {/* backend not ready yet — ignore */})
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  // ── Handlers ───────────────────────────────────────────────────────────


  const handleTypeChange = async (type: CertType) => {
    setActiveCertType(type)
    await setActiveType(type)
    window.location.reload()
  }

  const handleExcelFile = useCallback(async (file: File) => {

    setExcelLoading(true); setExcelError('')
    try {
      const result = await uploadExcel(file)
      // Set the excel result directly from the upload response — columns come from the file itself
      setExcelResult(result)
      // Reset column mapping so user maps fresh for the new file
      setColMap({ student_name: '', school: '', event_name: '', poc_email: '' })
      setColGroups([])
    } catch (e: any) {
      setExcelError(e.message)
    } finally {
      setExcelLoading(false)
    }
  }, [])

  const handleClearExcel = useCallback(async () => {
    try {
      await clearExcel()
      setExcelResult(null)
      setColMap({ student_name: '', school: '', event_name: '', poc_email: '' })
      setColGroups([])
    } catch {
      // ignore
    }
  }, [])

  /**
   * Loads the preview image + text fields from backend.
   * Exposed separately so it can be called both from the button
   * and automatically after a successful template upload.
   */
  const handleGridPreview = useCallback(async (posOverride?: '1st' | '2nd' | '3rd') => {
    setGridLoading(true)
    setPreviewError('')
    try {
      const pos = posOverride || activeWinnerPosRef.current
      const res = await getCoordinatesPreview(pos)
      setGridImage(res.image)
      setGridWidth(res.width || 1200)
      setGridHeight(res.height || 850)
      const tf = await getTextFields(pos)
      const enriched = Object.fromEntries(
        Object.entries(tf.image_text_fields).map(([k, v]: [string, any]) => [
          k,
          { ...v, font_family: v.font_family || "'DejaVu Sans', sans-serif" }
        ])
      )
      setTextFields(enriched)
      const keys = Object.keys(enriched)
      if (keys.length > 0) setActiveField(keys[0])
    } catch (e: any) {
      setPreviewError(
        e?.message ||
        'Could not generate preview. Make sure the backend is running and an image/PDF template is uploaded.'
      )
    } finally {
      setGridLoading(false)
    }
  }, [])

  const handleTemplateFile = useCallback(async (file: File, winnerPosition?: '1st' | '2nd' | '3rd') => {
    setTemplateLoading(true); setTemplateError('')
    try {
      const res = await uploadTemplate(file, winnerPosition)
      if (winnerPosition === '1st') setTemplatePath1st(res.path)
      else if (winnerPosition === '2nd') setTemplatePath2nd(res.path)
      else if (winnerPosition === '3rd') setTemplatePath3rd(res.path)
      
      if (winnerPosition) {
        setActiveWinnerPos(winnerPosition)
        activeWinnerPosRef.current = winnerPosition
      }

      setTemplatePath(res.path)
      const mode = res.template_mode as 'docx' | 'image'
      setTemplateMode(mode)
      // ── Auto-load preview immediately when an image template is uploaded ──
      if (mode === 'image') {
        // Small delay so the backend has finished writing the file
        await new Promise(r => setTimeout(r, 300))
        await handleGridPreview(winnerPosition)
      }
    } catch (e: any) {
      setTemplateError(e.message)
    } finally {
      setTemplateLoading(false)
    }
  }, [handleGridPreview])

  const handleSaveColumns = async () => {
    setColSaving(true); setColSaved(false)
    try {
      const res = await saveColumns(colMap)
      setColGroups(res.groups)
      setColSaved(true)
    } finally {
      setColSaving(false)
    }
  }

  // ── Save helpers ───────────────────────────────────────────────────────

  const handleSaveFields = useCallback(async (fields?: Record<string, any>) => {
    const toSave = fields ?? textFieldsRef.current
    await saveTextFields(toSave, activeWinnerPosRef.current)
    setFieldsSaved(true)
  }, [])

  const updateField = (name: string, key: string, value: any) => {
    setTextFields(tf => ({ ...tf, [name]: { ...tf[name], [key]: value } }))
    setFieldsSaved(false)
  }

  // ── Snap engine ────────────────────────────────────────────────────────

  /**
   * Compute all snap axes for a drag gesture and cache them in the ref.
   * Called ONCE on pointerdown — never during move.
   *
   * Axes contributed:
   *   • Canvas center + quarter lines (25%, 50%, 75%) on both axes
   *   • Each sibling field: anchor x and y
   *
   * We intentionally keep it anchor-to-anchor here because we don't know
   * the rendered text width without a full layout pass. The snapping still
   * gives accurate "align to the same column/row as another field" behaviour.
   */
  const computeSnapTargets = useCallback((draggingFieldName: string) => {
    const { w, h } = gridDimsRef.current
    const xAxes: SnapAxis[] = []
    const yAxes: SnapAxis[] = []

    // ── Canvas geometry anchors ──
    for (const frac of [0.25, 0.5, 0.75]) {
      xAxes.push({ canvasValue: w * frac, pct: frac * 100 })
      yAxes.push({ canvasValue: h * frac, pct: frac * 100 })
    }
    // Also snap to exact left / right / top / bottom edges
    xAxes.push({ canvasValue: 0, pct: 0 }, { canvasValue: w, pct: 100 })
    yAxes.push({ canvasValue: 0, pct: 0 }, { canvasValue: h, pct: 100 })

    // ── Object-to-object snap targets ──
    // For each sibling field contribute its anchor + estimated edges
    const fields = textFieldsRef.current
    Object.entries(fields).forEach(([name, cfg]: [string, any]) => {
      if (name === draggingFieldName) return
      if (cfg.x === undefined || cfg.y === undefined) return

      const fx: number = cfg.x
      const fy: number = cfg.y

      // Anchor point (the defined x, y)
      xAxes.push({ canvasValue: fx, pct: (fx / w) * 100 })
      yAxes.push({ canvasValue: fy, pct: (fy / h) * 100 })

      // Estimated half-width of the text element so we can offer
      // left-edge / right-edge alignment as well.
      // Rough heuristic: font_size * 0.55 * ~8 chars average = 4.4 * font_size
      // capped to a reasonable max.
      const estimatedHalfW = Math.min((cfg.font_size || 24) * 3.5, w * 0.25)
      const estimatedHalfH = (cfg.font_size || 24) * 0.6

      const leftEdge  = fx - estimatedHalfW
      const rightEdge = fx + estimatedHalfW
      const topEdge   = fy - estimatedHalfH
      const botEdge   = fy + estimatedHalfH

      if (leftEdge > 0)  xAxes.push({ canvasValue: leftEdge,  pct: (leftEdge  / w) * 100 })
      if (rightEdge < w) xAxes.push({ canvasValue: rightEdge, pct: (rightEdge / w) * 100 })
      if (topEdge > 0)   yAxes.push({ canvasValue: topEdge,   pct: (topEdge   / h) * 100 })
      if (botEdge < h)   yAxes.push({ canvasValue: botEdge,   pct: (botEdge   / h) * 100 })
    })

    // Deduplicate axes that are within 2px of each other (prevents double-firing)
    const dedup = (axes: SnapAxis[]): SnapAxis[] => {
      const sorted = [...axes].sort((a, b) => a.canvasValue - b.canvasValue)
      return sorted.filter((ax, i) =>
        i === 0 || Math.abs(ax.canvasValue - sorted[i - 1].canvasValue) > 2
      )
    }

    snapTargetsRef.current = { x: dedup(xAxes), y: dedup(yAxes) }
  }, [])

  /**
   * Apply snap: given a raw canvas-coordinate value and a list of axes,
   * return the snapped value + any triggered guide percentages.
   * Guide triggers AND holds while distance ≤ SNAP_THRESHOLD (5px).
   * Dissolves immediately when distance > threshold.
   */
  const applySnap = (
    rawValue: number,
    axes: SnapAxis[],
  ): { snapped: number; guides: number[] } => {
    let best: SnapAxis | null = null
    let bestDist = Infinity
    for (const ax of axes) {
      const d = Math.abs(rawValue - ax.canvasValue)
      if (d <= SNAP_THRESHOLD && d < bestDist) {
        best = ax
        bestDist = d
      }
    }
    if (best) {
      return { snapped: best.canvasValue, guides: [best.pct] }
    }
    return { snapped: rawValue, guides: [] }
  }

  // ── Pointer / mouse handlers ───────────────────────────────────────────

  /**
   * onMouseDown for each draggable field element.
   * Does NOT use setPointerCapture — that would redirect pointermove events
   * away from the canvas container where onMouseMove is attached.
   * Instead we store the dragging field in a ref (always-current) and let
   * the canvas container's onMouseMove do all the work.
   */
  const handleFieldMouseDown = useCallback((
    e: React.MouseEvent<HTMLDivElement>,
    name: string,
  ) => {
    e.preventDefault()
    e.stopPropagation()
    draggingFieldRef.current = name   // sync, never stale
    setDraggingField(name)            // triggers re-render for cursor / ring
    setActiveField(name)
    // Pre-compute snap targets once for this gesture
    computeSnapTargets(name)
  }, [computeSnapTargets])

  /**
   * onMouseMove on the canvas container.
   * Reads draggingFieldRef (not state) so it is never a stale closure,
   * regardless of useCallback dependency arrays.
   * All arithmetic is in canvas-space (original image pixel coordinates).
   */
  const handleMouseMove = useCallback((e: React.MouseEvent<HTMLDivElement>) => {
    const rect = e.currentTarget.getBoundingClientRect()
    const { w, h } = gridDimsRef.current

    // Convert cursor to canvas-space coords (clamped to image bounds)
    const cursorX = Math.max(0, Math.min(rect.width,  e.clientX - rect.left))
    const cursorY = Math.max(0, Math.min(rect.height, e.clientY - rect.top))
    const rawCanvasX = (cursorX / rect.width)  * w
    const rawCanvasY = (cursorY / rect.height) * h

    // ── Handle Line Drawing (read from refs — never stale) ──
    if (isDrawingLineRef.current && drawStartRef.current) {
      const fieldName = activeFieldRef.current
      if (!fieldName) return
      const startX = drawStartRef.current.x
      const startY = drawStartRef.current.y  // Y is locked to where click started
      const lineLeft  = Math.min(startX, rawCanvasX)
      const lineRight = Math.max(startX, rawCanvasX)

      // Update live visual preview line (CSS percentages)
      setDrawingPreview({
        x1Pct: (lineLeft  / w) * 100,
        x2Pct: (lineRight / w) * 100,
        yPct:  (startY    / h) * 100,
      })

      // Live-update the text field so the bounding-box overlay moves too
      setTextFields(tf => ({
        ...tf,
        [fieldName]: {
          ...tf[fieldName],
          x: Math.round(lineLeft),
          width: Math.round(lineRight - lineLeft),
          y: Math.round(startY),
        },
      }))
      setFieldsSaved(false)
      return
    }

    // ── Handle Normal Dragging ──
    const name = draggingFieldRef.current
    if (!name) return


    // Apply snap independently on each axis
    const { snapped: snappedX, guides: vGuides } = applySnap(rawCanvasX, snapTargetsRef.current.x)
    const { snapped: snappedY, guides: hGuides } = applySnap(rawCanvasY, snapTargetsRef.current.y)

    setGuides({ h: hGuides, v: vGuides })
    setTextFields(tf => ({
      ...tf,
      [name]: { ...tf[name], x: Math.round(snappedX), y: Math.round(snappedY) },
    }))
    setFieldsSaved(false)
  }, [])  // no deps — reads refs directly, never stale

  const handleCanvasMouseUp = useCallback(() => {
    if (isDrawingLineRef.current && drawStartRef.current) {
      // Finish this stroke: clear the drag anchor and preview line,
      // but KEEP isDrawingLine=true so user can draw again without re-clicking.
      drawStartRef.current = null
      setDrawingPreview(null)
      handleSaveFields()
      return
    }
    if (!draggingFieldRef.current) return   // nothing was being dragged
    draggingFieldRef.current = null
    setDraggingField(null)
    setGuides({ h: [], v: [] })
    handleSaveFields()
  }, [handleSaveFields])

  const handleCanvasMouseDown = useCallback((e: React.MouseEvent<HTMLDivElement>) => {
    // Use activeFieldRef — activeField state is stale in useCallback([])
    if (isDrawingLineRef.current && activeFieldRef.current) {
      const rect = e.currentTarget.getBoundingClientRect()
      const { w, h } = gridDimsRef.current
      const cursorX = Math.max(0, Math.min(rect.width,  e.clientX - rect.left))
      const cursorY = Math.max(0, Math.min(rect.height, e.clientY - rect.top))
      drawStartRef.current = {
        x: (cursorX / rect.width)  * w,
        y: (cursorY / rect.height) * h
      }
      e.stopPropagation()
    }
  }, [])  // no deps — reads refs, never stale

  // ── Auth handlers ──────────────────────────────────────────────────────

  const authReady = senderEmail.trim() !== '' && appPassword.trim() !== ''

  const handleSaveAuth = async () => {
    setAuthSaving(true); setAuthError(''); setAuthSaved(false)
    try {
      await setEmailAuth(senderEmail.trim(), appPassword.trim())
      setAuthSaved(true)
      setSavedEmail(senderEmail.trim())
      setAppPassword('')
    } catch (e: any) {
      setAuthError(e.message)
    } finally {
      setAuthSaving(false)
    }
  }

  const handleSignOut = async () => {
    setSigningOut(true)
    try {
      await signOutEmailAuth()
      setAuthSaved(false)
      setSavedEmail(null)
      setSenderEmail('')
      setAppPassword('')
    } finally {
      setSigningOut(false)
    }
  }

  const handleDriveFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0]
    if (!file) return
    setDriveUploading(true)
    setDriveError('')
    try {
      await uploadDriveCredentials(file)
      const status = await getDriveAuthStatus()
      setDriveAvailable(status.drive_available)
      setDriveCredsExist(status.has_credentials)
    } catch (err: any) {
      setDriveError(err.message || 'Failed to upload service_account.json')
    } finally {
      setDriveUploading(false)
    }
  }

  const handleDeleteDrive = async () => {
    setDriveUploading(true)
    try {
      await deleteDriveCredentials()
      setDriveAvailable(false)
      setDriveCredsExist(false)
      setServiceAccountEmail(null)
    } finally {
      setDriveUploading(false)
    }
  }

  const handleSaveDriveFolder = async () => {
    setSaveFolderLoading(true)
    try {
      await saveDriveRootFolder(driveFolderInput.trim())
      setDriveFolderSaved(true)
      setTimeout(() => setDriveFolderSaved(false), 3000)
    } catch (err: any) {
      setDriveError(err.message || 'Failed to save Drive root folder')
    } finally {
      setSaveFolderLoading(false)
    }
  }

  const handleOAuthClientUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0]
    if (!file) return
    setOauthUploading(true)
    setDriveError('')
    try {
      await uploadOAuthClientSecrets(file)
      setOauthClientAvailable(true)
    } catch (err: any) {
      setDriveError(err.message || 'Failed to upload OAuth2 client_secrets.json')
    } finally {
      setOauthUploading(false)
    }
  }

  const handleStartOAuth = async () => {
    setOauthFlowStatus('in_progress')
    setOauthFlowError('')
    try {
      // Build the redirect URI — this is the backend /oauth-callback endpoint
      const redirectUri = `${getBackendUrl()}/config/drive-auth/oauth-callback`
      const { url } = await getOAuthUrl(redirectUri)
      // Open Google's consent page in a new tab (works in cloud deployments)
      window.open(url, '_blank', 'noopener,noreferrer')

      // Poll the backend until the callback completes the flow
      const poll = setInterval(async () => {
        try {
          const s = await getOAuthFlowStatus()
          if (s.status === 'done') {
            clearInterval(poll)
            setOauthFlowStatus('done')
            setOauthDriveAvailable(true)
            setOauthUserEmail(s.oauth_user_email)
            setDriveAvailable(true)
          } else if (s.status === 'error') {
            clearInterval(poll)
            setOauthFlowStatus('error')
            setOauthFlowError(s.error || 'Authorization failed')
          }
        } catch { /* keep polling */ }
      }, 1500)
    } catch (err: any) {
      setOauthFlowStatus('error')
      setOauthFlowError(err.message || 'Failed to start OAuth2 flow')
    }
  }

  const handleRevokeOAuth = async () => {
    try {
      await revokeOAuthToken()
      setOauthDriveAvailable(false)
      setOauthUserEmail(null)
      setOauthFlowStatus('idle')
      if (!driveCredsExist) setDriveAvailable(false)
    } catch (err: any) {
      setDriveError(err.message || 'Failed to revoke OAuth2 token')
    }
  }

  // ── Scale helper ───────────────────────────────────────────────────────

  /** Ratio of rendered image width to original canvas width */
  const getScale = () => {
    if (!imgRef.current) return 1
    return imgRef.current.clientWidth / gridWidth
  }

  const displayGroups = colGroups.length ? colGroups : (excelResult?.groups ?? [])

  // ── Render ─────────────────────────────────────────────────────────────

  // ── Setup Wizard state ─────────────────────────────────────────────────
  const [wizardDismissed, setWizardDismissed] = useState(() => localStorage.getItem('wizard_dismissed') === '1')
  const wizardSteps = [
    { label: 'Upload roster Excel', done: !!excelResult, hint: 'Roster Data section below' },
    { label: 'Map Excel columns', done: colSaved, hint: 'Column Mapping section below' },
    { label: 'Upload certificate template', done: !!(templatePath || templatePath1st), hint: 'Certificate Template section below' },
    { label: 'Set up Gmail (SMTP)', done: !!savedEmail, hint: 'Email Credentials section below' },
    { label: 'Google Drive (optional)', done: oauthDriveAvailable || driveAvailable, hint: 'Google Drive section below', optional: true },
  ]
  const requiredDone = wizardSteps.filter(s => !s.optional).every(s => s.done)

  return (

    <div>
      {/* ── Setup Wizard ── */}
      {!wizardDismissed && (
        <div className="card" style={{ marginBottom: 24, borderColor: requiredDone ? '#10b981' : '#6366f1', borderLeft: '4px solid ' + (requiredDone ? '#10b981' : '#6366f1'), background: requiredDone ? '#f0fdf4' : undefined }}>
          <div className="card-header" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', background: requiredDone ? '#dcfce7' : undefined, color: requiredDone ? '#15803d' : undefined }}>
            <span>{requiredDone ? '🎉 All set! Ready to generate & send certificates.' : '🚀 Setup Checklist — complete these steps to get started'}</span>
            <button
              onClick={() => { setWizardDismissed(true); localStorage.setItem('wizard_dismissed', '1') }}
              style={{ background: 'none', border: 'none', cursor: 'pointer', color: '#94a3b8', fontSize: 18, lineHeight: 1 }}
              title="Dismiss"
            >✕</button>
          </div>
          <div className="card-body" style={{ padding: '12px 20px' }}>
            <div style={{ display: 'flex', flexWrap: 'wrap', gap: 10 }}>
              {wizardSteps.map((s, i) => (
                <div key={i} style={{
                  display: 'flex', alignItems: 'center', gap: 8,
                  padding: '8px 14px', borderRadius: 8, fontSize: 13, fontWeight: 500,
                  background: s.done ? '#dcfce7' : s.optional ? '#f8fafc' : '#eff6ff',
                  color: s.done ? '#15803d' : s.optional ? '#94a3b8' : '#1e40af',
                  border: '1px solid ' + (s.done ? '#86efac' : s.optional ? '#e2e8f0' : '#bfdbfe'),
                }}>
                  <span style={{ fontSize: 16 }}>{s.done ? '✅' : s.optional ? '⭕' : '⏳'}</span>
                  <span>{i + 1}. {s.label}{s.optional ? ' (optional)' : ''}</span>
                </div>
              ))}
            </div>
            {!requiredDone && (
              <div style={{ marginTop: 10, fontSize: 12, color: '#64748b' }}>
                Complete all required steps above, then go to the <strong>Generate &amp; Send</strong> tab.
              </div>
            )}
          </div>
        </div>
      )}

      {/* ── Certificate Type Selector ── */}
      <div className="card" style={{ marginBottom: 24, borderColor: '#3b82f6', borderLeft: '4px solid #3b82f6' }}>
        <div className="card-header" style={{ backgroundColor: '#eff6ff', color: '#1e40af' }}>📝 Step 1: Select Certificate Type</div>
        <div className="card-body">
          <div style={{ display: 'flex', gap: 20 }}>
            {CERT_TYPES.map(t => (
              <label key={t.id} style={{ display: 'flex', alignItems: 'center', gap: 8, cursor: 'pointer', fontWeight: activeCertType === t.id ? 600 : 400 }}>
                <input 
                  type="radio" 
                  name="cert_type" 
                  checked={activeCertType === t.id} 
                  onChange={() => handleTypeChange(t.id)} 
                  style={{ transform: 'scale(1.2)' }}
                />
                {t.label}
              </label>
            ))}
          </div>
        </div>
      </div>

      {/* ── Excel Upload ── */}

      <div className="card">
        <div className="card-header"><FileSpreadsheet size={18} /> Roster Data (Excel)</div>
        <div className="card-body">
          <input type="file" accept=".xlsx, .xls" onChange={e => { const f = e.target.files?.[0]; if (f) handleExcelFile(f); e.target.value = '' }} />
          {excelLoading && <div className="alert alert-info mt-2" style={{ fontSize: 13 }}>⏳ Reading Excel file…</div>}
          {excelError && <div className="alert alert-danger mt-2" style={{ fontSize: 13 }}>❌ {excelError}</div>}
          {excelResult && (
            <div className="alert alert-success mt-2" style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
              <CheckCircle2 size={18} />
              <span>Uploaded: <strong>{excelResult.excel_path.split('/').pop()}</strong> — {excelResult.excel_columns.length} columns detected</span>
              <button
                onClick={handleClearExcel}
                style={{ marginLeft: 'auto', background: 'none', border: '1px solid #dc2626', color: '#dc2626', borderRadius: 6, padding: '2px 10px', cursor: 'pointer', fontSize: 12 }}
                title="Remove this Excel file and start fresh"
              >✕ Remove</button>
            </div>
          )}
          {displayGroups.length > 0 && (
            <div className="mt-3">
              <div className="section-title">Roster Preview — grouped by POC email</div>
              {displayGroups.map((g, i) => (
                <div key={i} className="card" style={{ marginBottom: 16 }}>
                  <div className="card-header" style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
                    <Mail size={16} className="text-muted" /> {g.poc_name ? `${g.poc_name} (${g.poc_email})` : g.poc_email}
                    <span className="badge-success">{g.student_count} student{g.student_count !== 1 ? 's' : ''}</span>
                  </div>
                  <div className="card-body">
                    <table className="data-table">
                      <thead>
                        <tr>
                          <th>#</th>
                          <th>Name</th>
                          {activeCertType === 'winner' && <th>Position</th>}
                          <th>School</th>
                          <th>Event</th>
                        </tr>
                      </thead>
                      <tbody>
                        {(g.student_details ?? g.students.map(s => ({ name: s, school: '', event_name: '', position: '' }))).map((sd, j) => (
                          <tr key={j}>
                            <td>{j + 1}</td>
                            <td>{sd.name}</td>
                            {activeCertType === 'winner' && <td>{sd.position || ''}</td>}
                            <td>{sd.school}</td>
                            <td>{sd.event_name}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              ))}
            </div>
          )}
        </div>
      </div>

      {/* ── Column Mapping ── */}
      <div className="card">
        <div className="card-header">🔗 Column Mapping</div>
        <div className="card-body">
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 16 }}>
            {getKeysForType(activeCertType).map(key => (
              <div className="form-group" key={key}>
                <label>
                  {KEY_LABELS[key as InternalKey]}
                  <span className="text-muted" style={{ fontSize: 11, marginLeft: 6, fontWeight: 400 }}>
                    {KEY_HINTS[key as InternalKey]}
                  </span>
                </label>
                <select
                  value={colMap[key] ?? ''}
                  onChange={e => setColMap(m => ({ ...m, [key]: e.target.value }))}
                >
                  <option value="">-- Select Column --</option>
                  {excelResult?.excel_columns.map(c => (
                    <option key={c} value={c}>{c}</option>
                  ))}
                </select>
              </div>
            ))}
          </div>
          <div style={{ marginTop: 24, display: 'flex', alignItems: 'center', gap: 12 }}>
            <button className="btn btn-primary" onClick={handleSaveColumns} disabled={colSaving}>
              {colSaving ? '…Saving' : <><Save size={16} /> Save Mapping</>}
            </button>
            {colSaved && <span className="badge-success"><CheckCircle2 size={14} /> Saved</span>}
          </div>
        </div>
      </div>

      {/* ── Certificate Template ── */}
      <div className="card">
        <div className="card-header"><FileImage size={18} /> Certificate Template</div>
        <div className="card-body">
          <div className="flex gap-4 mb-4">
            <label><input type="radio" checked={templateMode === 'docx'} onChange={() => setTemplateMode('docx')} /> Word (.docx)</label>
            <label><input type="radio" checked={templateMode === 'image'} onChange={() => setTemplateMode('image')} /> Image / PDF</label>
          </div>
          {activeCertType === 'winner' && templateMode === 'image' ? (
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr', gap: 16, marginTop: 12 }}>
              {/* 1st Place */}
              <div style={{ border: '1px dashed #ccc', borderRadius: 6, padding: 12 }}>
                <div style={{ fontWeight: 600, fontSize: 13, marginBottom: 8 }}>🥇 1st Place Certificate</div>
                <input
                  type="file"
                  onChange={e => { const f = e.target.files?.[0]; if (f) handleTemplateFile(f, '1st') }}
                  style={{ fontSize: 12, width: '100%' }}
                />
                {templatePath1st && (
                  <div style={{ fontSize: 11, color: '#16a34a', marginTop: 6, wordBreak: 'break-all' }}>
                    ✔ {templatePath1st.split('/').pop()}
                  </div>
                )}
              </div>

              {/* 2nd Place */}
              <div style={{ border: '1px dashed #ccc', borderRadius: 6, padding: 12 }}>
                <div style={{ fontWeight: 600, fontSize: 13, marginBottom: 8 }}>🥈 2nd Place Certificate</div>
                <input
                  type="file"
                  onChange={e => { const f = e.target.files?.[0]; if (f) handleTemplateFile(f, '2nd') }}
                  style={{ fontSize: 12, width: '100%' }}
                />
                {templatePath2nd && (
                  <div style={{ fontSize: 11, color: '#16a34a', marginTop: 6, wordBreak: 'break-all' }}>
                    ✔ {templatePath2nd.split('/').pop()}
                  </div>
                )}
              </div>

              {/* 3rd Place */}
              <div style={{ border: '1px dashed #ccc', borderRadius: 6, padding: 12 }}>
                <div style={{ fontWeight: 600, fontSize: 13, marginBottom: 8 }}>🥉 3rd Place Certificate</div>
                <input
                  type="file"
                  onChange={e => { const f = e.target.files?.[0]; if (f) handleTemplateFile(f, '3rd') }}
                  style={{ fontSize: 12, width: '100%' }}
                />
                {templatePath3rd && (
                  <div style={{ fontSize: 11, color: '#16a34a', marginTop: 6, wordBreak: 'break-all' }}>
                    ✔ {templatePath3rd.split('/').pop()}
                  </div>
                )}
              </div>
            </div>
          ) : (
            <input
              type="file"
              onChange={e => { const f = e.target.files?.[0]; if (f) handleTemplateFile(f) }}
            />
          )}

          {templateLoading && (
            <div className="alert alert-info mt-2" style={{ fontSize: 13 }}>
              ⏳ Uploading and generating preview…
            </div>
          )}

          {templatePath && !templateLoading && activeCertType !== 'winner' && (
            <div className="alert alert-success mt-2">
              <CheckCircle2 size={18} /> {templatePath.split('/').pop()}
              {templateMode === 'image' && gridImage && (
                <span style={{ marginLeft: 8, fontSize: 12, opacity: 0.8 }}>— Preview auto-loaded ↓</span>
              )}
            </div>
          )}

          {activeCertType === 'winner' && templateMode === 'image' && templatePath && !templateLoading && (
            <div className="alert alert-success mt-2" style={{ fontSize: 12 }}>
              <CheckCircle2 size={16} style={{ marginRight: 8, display: 'inline', verticalAlign: 'middle' }} />
              You can drag and drop text attributes independently for 1st, 2nd, and 3rd place certificates using the position tabs in the Visual Editor below.
            </div>
          )}
        </div>
      </div>

      {/* ── Visual Editor (image mode only) ── */}
      {templateMode === 'image' && (
        <div className="card">
          <div className="card-header"><Type size={18} /> Visual Editor</div>
          <div className="card-body">
            {activeCertType === 'winner' && (
              <div style={{ display: 'flex', alignItems: 'center', gap: 12, marginBottom: 16, background: '#f8fafc', padding: '10px 14px', borderRadius: 8, border: '1px solid #e2e8f0' }}>
                <span style={{ fontWeight: 600, fontSize: 13, color: '#334155' }}>🎯 Visual Editing Target:</span>
                {(['1st', '2nd', '3rd'] as const).map(pos => (
                  <button
                    key={pos}
                    className={`btn btn-sm ${activeWinnerPos === pos ? 'btn-primary' : 'btn-ghost'}`}
                    style={{ border: activeWinnerPos === pos ? 'none' : '1px solid #cbd5e1', fontWeight: 600 }}
                    onClick={() => {
                      setActiveWinnerPos(pos)
                      handleGridPreview(pos)
                    }}
                  >
                    {pos === '1st' ? '🥇 1st Place' : pos === '2nd' ? '🥈 2nd Place' : '🥉 3rd Place'}
                  </button>
                ))}
              </div>
            )}

            <button className="btn btn-ghost" onClick={() => handleGridPreview()} disabled={gridLoading}>
              {gridLoading ? '…Loading' : '🔄 Refresh Preview'}
            </button>
            <p className="text-muted mt-2" style={{ marginBottom: 12, fontSize: 13 }}>
              Drag any text placeholder freely. Smart guides (<span style={{ color: SNAP_COLOR, fontWeight: 600 }}>magenta lines</span>) snap to canvas axes and other field edges within 5 px — dissolve automatically on departure.
            </p>

            {previewError && (
              <div className="alert alert-danger mt-2" style={{ marginBottom: 12 }}>
                ⚠️ {previewError}
              </div>
            )}

            {gridImage && (
              <>
                {/* ── Canvas container ── */}
                <div
                  ref={canvasContainerRef}
                  style={{
                    position: 'relative',
                    display: 'inline-block',
                    width: '100%',
                    cursor: isDrawingLine ? 'crosshair' : draggingField ? 'grabbing' : 'default',
                    marginBottom: 20,
                    userSelect: 'none',
                    borderRadius: 8,
                    overflow: 'hidden',
                    boxShadow: '0 4px 24px rgba(0,0,0,0.3)',
                  }}
                  onMouseDown={handleCanvasMouseDown}
                  onMouseMove={handleMouseMove}
                  onMouseUp={handleCanvasMouseUp}
                  onMouseLeave={handleCanvasMouseUp}
                >
                  <img
                    ref={imgRef}
                    src={gridImage}
                    alt="Certificate preview"
                    style={{ width: '100%', display: 'block' }}
                    draggable={false}
                  />

                  {/* ── Smart Guide Lines Overlay ── */}
                  {guides.v.map((vPct, i) => (
                    <div key={`v-${i}`} style={{
                      position: 'absolute',
                      top: 0, bottom: 0,
                      left: `${vPct}%`,
                      width: 1,
                      background: SNAP_COLOR,
                      boxShadow: `0 0 4px ${SNAP_COLOR}, 0 0 8px ${SNAP_COLOR}40`,
                      pointerEvents: 'none',
                      zIndex: 20,
                    }} />
                  ))}
                  {guides.h.map((hPct, i) => (
                    <div key={`h-${i}`} style={{
                      position: 'absolute',
                      left: 0, right: 0,
                      top: `${hPct}%`,
                      height: 1,
                      background: SNAP_COLOR,
                      boxShadow: `0 0 4px ${SNAP_COLOR}, 0 0 8px ${SNAP_COLOR}40`,
                      pointerEvents: 'none',
                      zIndex: 20,
                    }} />
                  ))}

                  {/* ── Live Baseline Drawing Preview Line ── */}
                  {drawingPreview && (
                    <div
                      style={{
                        position: 'absolute',
                        top: `${drawingPreview.yPct}%`,
                        left: `${drawingPreview.x1Pct}%`,
                        width: `${drawingPreview.x2Pct - drawingPreview.x1Pct}%`,
                        height: 3,
                        background: 'rgba(99, 230, 190, 0.95)',
                        boxShadow: '0 0 8px rgba(99,230,190,0.8), 0 0 20px rgba(99,230,190,0.4)',
                        pointerEvents: 'none',
                        zIndex: 25,
                        borderRadius: 2,
                        transform: 'translateY(-50%)',
                      }}
                    />
                  )}

                  {/* ── Persisted Baseline Lines (for fields with a drawn width) ── */}
                  {Object.entries(textFields).map(([name, cfg]: [string, any]) => {
                    if (!cfg.width || cfg.width <= 0) return null
                    const x1Pct = (cfg.x / gridWidth) * 100
                    const x2Pct = ((cfg.x + cfg.width) / gridWidth) * 100
                    const yPct  = (cfg.y / gridHeight) * 100
                    const isActive = name === activeField
                    return (
                      <div
                        key={`baseline-${name}`}
                        style={{
                          position: 'absolute',
                          top: `${yPct}%`,
                          left: `${x1Pct}%`,
                          width: `${x2Pct - x1Pct}%`,
                          height: 2,
                          background: isActive ? 'rgba(99,230,190,0.9)' : 'rgba(99,230,190,0.45)',
                          boxShadow: isActive ? '0 0 6px rgba(99,230,190,0.7)' : 'none',
                          pointerEvents: 'none',
                          zIndex: 18,
                          borderRadius: 1,
                          transform: 'translateY(-50%)',
                        }}
                      />
                    )
                  })}

                  {/* ── Text Field Overlays ── */}
                  <div style={{ position: 'absolute', inset: 0, pointerEvents: 'none' }}>
                    {Object.entries(textFields).map(([name, cfg]: [string, any]) => {
                      if (cfg.x === undefined || cfg.y === undefined) return null

                      const pctX = (cfg.x / gridWidth) * 100
                      const pctY = (cfg.y / gridHeight) * 100
                      const isActive = name === activeField
                      const isDragging = name === draggingField
                      const scale = getScale()

                      const transformX =
                        cfg.align === 'center' ? '-50%' :
                        cfg.align === 'right'  ? '-100%' :
                        '0%'

                      const fontFamily = cfg.font_family || "'DejaVu Sans', sans-serif"
                      const previewText =
                        name === 'student_name' ? 'Jane Doe' :
                        name === 'volunteer_name' ? 'John Smith' :
                        name === 'school'        ? 'Christ University' :
                        name === 'event_name'    ? 'TechFest 2026' :
                        name

                      return (
                        <div
                          key={name}
                          data-snap-field={name}
                          onMouseDown={e => handleFieldMouseDown(e, name)}
                          style={{
                            position: 'absolute',
                            left: `${pctX}%`,
                            top: `${pctY}%`,
                            transform: cfg.width > 0 ? 'translate(0, -85%)' : `translate(${transformX}, -50%)`,
                            cursor: isDrawingLine ? 'crosshair' : isDragging ? 'grabbing' : 'grab',
                            zIndex: isActive ? 15 : 5,
                            pointerEvents: 'auto',
                            color: cfg.color || '#ffffff',
                            fontSize: `${(cfg.font_size || 24) * scale}px`,
                            fontWeight: cfg.is_bold ? 'bold' : 'normal',
                            fontFamily,
                            whiteSpace: cfg.width > 0 ? 'normal' : 'nowrap',
                            lineHeight: 1,
                            userSelect: 'none',
                            opacity: isDragging ? 0.85 : 1,
                            textShadow: '0 1px 8px rgba(0,0,0,0.95), 0 0 2px rgba(0,0,0,0.8)',
                            padding: cfg.width > 0 ? '0px 6px' : '2px 6px',
                            width: cfg.width ? `${(cfg.width / gridWidth) * 100}%` : undefined,
                            textAlign: cfg.width ? 'center' : undefined,
                            outline: isDragging
                              ? `1.5px solid ${SNAP_COLOR}`
                              : isActive
                              ? '1.5px dashed rgba(250,204,21,0.85)'
                              : '1px dashed rgba(255,255,255,0.35)',
                            outlineOffset: 4,
                            borderRadius: 2,
                            transition: isDragging ? 'none' : 'outline 0.12s ease',
                            filter: isDragging
                              ? `drop-shadow(0 0 6px ${SNAP_COLOR}80)`
                              : isActive
                              ? 'drop-shadow(0 0 4px rgba(250,204,21,0.5))'
                              : undefined,
                          }}
                        >
                          {previewText}
                        </div>
                      )
                    })}

                  </div>
                </div>
              </>
            )}

            {/* ── Field Properties Table ── */}


            {Object.keys(textFields).length > 0 && (
              <div className="mt-3">
                <div className="section-title" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                  <div>
                    Text Fields {activeField && (
                      <span className="badge-success" style={{ marginLeft: 8, fontSize: 11 }}>
                        Active: {activeField}
                      </span>
                    )}
                  </div>
                  {activeField && (
                    <button
                      className={`btn btn-sm ${isDrawingLine ? 'btn-primary' : 'btn-ghost'}`}
                      onClick={() => {
                        const next = !isDrawingLine
                        setIsDrawingLine(next)
                        isDrawingLineRef.current = next
                      }}
                      style={{ border: '1px solid var(--primary)', color: isDrawingLine ? '#fff' : 'var(--primary)' }}
                    >
                      ✏️ {isDrawingLine ? 'Cancel Drawing' : 'Draw Baseline'}
                    </button>
                  )}
                </div>
                <div style={{ overflowX: 'auto' }}>
                  <table className="data-table" style={{ marginBottom: 16, minWidth: 720 }}>
                    <thead>
                      <tr>
                        <th>Field</th>
                        <th>X</th>
                        <th>Y</th>
                        <th>Width</th>
                        <th>Size</th>
                        <th>Color</th>
                        <th>Font</th>
                        <th>Weight</th>
                        <th>Align</th>
                      </tr>
                    </thead>
                    <tbody>
                      {Object.entries(textFields).map(([name, cfg]: [string, any]) => {
                        const isActive = name === activeField
                        const selectedFont = FONT_LIBRARY.find(f => f.serverPath === cfg.font_path) || FONT_LIBRARY[0]
                        return (
                          <tr
                            key={name}
                            onClick={() => setActiveField(name)}
                            style={{
                              cursor: 'pointer',
                              backgroundColor: isActive ? 'var(--primary-light)' : undefined,
                              borderLeft: isActive ? '4px solid var(--primary)' : '4px solid transparent'
                            }}
                          >
                            <td>
                              <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                                {isActive && <Move size={14} color="var(--primary)" />}
                                <strong style={{ color: isActive ? 'var(--primary)' : 'inherit' }}>{name}</strong>
                              </div>
                            </td>
                            {(['x', 'y', 'width', 'font_size'] as const).map(k => (
                              <td key={k} onClick={e => e.stopPropagation()}>
                                <input
                                  type="number"
                                  value={cfg[k] ?? 0}
                                  style={{ width: k === 'width' ? 80 : 70 }}
                                  placeholder={k === 'width' ? 'Auto' : ''}
                                  onChange={e => updateField(name, k, +e.target.value)}
                                />
                              </td>
                            ))}
                            <td onClick={e => e.stopPropagation()}>
                              <input
                                type="color"
                                value={cfg.color ?? '#000000'}
                                style={{ width: 36, height: 28, border: 'none', cursor: 'pointer', padding: 0 }}
                                onChange={e => updateField(name, 'color', e.target.value)}
                              />
                            </td>
                            {/* Font picker with live preview */}
                            <td onClick={e => e.stopPropagation()} style={{ minWidth: 180 }}>
                              <select
                                value={cfg.font_path ?? 'sample_data/DejaVuSans.ttf'}
                                onChange={e => {
                                  const chosen = FONT_LIBRARY.find(f => f.serverPath === e.target.value) || FONT_LIBRARY[0]
                                  setTextFields(tf => ({
                                    ...tf,
                                    [name]: {
                                      ...tf[name],
                                      font_path: chosen.serverPath,
                                      font_family: chosen.family,
                                    }
                                  }))
                                  setFieldsSaved(false)
                                }}
                                style={{ fontFamily: selectedFont.family, minWidth: 170 }}
                              >
                                {FONT_LIBRARY.map(f => (
                                  <option key={f.label} value={f.serverPath} style={{ fontFamily: f.family }}>
                                    {f.label}
                                  </option>
                                ))}
                              </select>
                              {/* Live preview swatch */}
                              <div style={{
                                fontFamily: selectedFont.family,
                                fontSize: 14,
                                color: cfg.color || '#000',
                                marginTop: 4,
                                background: 'rgba(128,128,128,0.1)',
                                padding: '2px 6px',
                                borderRadius: 4,
                                whiteSpace: 'nowrap',
                                overflow: 'hidden',
                                textOverflow: 'ellipsis',
                              }}>
                                {name === 'student_name' ? 'Amal Tom' : name === 'volunteer_name' ? 'John Smith' : name === 'event_name' ? 'Group Dance' : name === 'school' ? 'Christ University' : 'Sample'}
                              </div>
                            </td>
                            <td onClick={e => e.stopPropagation()}>
                              <select
                                value={cfg.is_bold ? 'bold' : 'normal'}
                                onChange={e => {
                                  const is_bold = e.target.value === 'bold'
                                  setTextFields(tf => ({
                                    ...tf,
                                    [name]: {
                                      ...tf[name],
                                      is_bold,
                                      font_path: is_bold ? 'sample_data/DejaVuSans-Bold.ttf' : 'sample_data/DejaVuSans.ttf',
                                    }
                                  }))
                                  setFieldsSaved(false)
                                }}
                              >
                                <option value="normal">Normal</option>
                                <option value="bold">Bold</option>
                              </select>
                            </td>
                            <td onClick={e => e.stopPropagation()}>
                              <select
                                value={cfg.align ?? 'left'}
                                onChange={e => updateField(name, 'align', e.target.value)}
                              >
                                <option value="left">Left</option>
                                <option value="center">Center</option>
                                <option value="right">Right</option>
                              </select>
                            </td>
                          </tr>
                        )
                      })}
                    </tbody>
                  </table>
                </div>
                <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
                  <button
                    className="btn btn-primary"
                    onClick={() => handleSaveFields()}
                    disabled={fieldsSaved}
                  >
                    <Save size={16} /> {fieldsSaved ? 'Saved ✓' : 'Save Text Fields'}
                  </button>
                  {fieldsSaved && <span className="badge-success"><CheckCircle2 size={14} /> Saved</span>}
                  
                  <button
                    className="btn btn-ghost btn-sm"
                    onClick={async () => {
                      if (!confirm("Reset all field positions to defaults?")) return;
                      await fetch(`${BASE}/config/reset-text-fields`, { method: 'POST' }).catch(() => {});
                      window.location.reload();
                    }}
                    style={{ marginLeft: 'auto', border: '1px solid #dc2626', color: '#dc2626' }}
                  >
                    Reset Positions
                  </button>
                </div>
              </div>
            )}
          </div>
        </div>
      )}

      {/* ── Auth Panel ── */}
      <div className="card">
        <div className="card-header"><Lock size={18} /> Gmail Account</div>
        <div className="card-body">
          {authSaved && savedEmail ? (
            <div style={{ display: 'flex', alignItems: 'center', gap: 16, flexWrap: 'wrap' }}>
              <div style={{
                display: 'flex', alignItems: 'center', gap: 10,
                background: 'var(--success-bg)',
                border: '1px solid var(--success-text)',
                borderRadius: 8, padding: '10px 16px', flex: 1,
              }}>
                <CheckCircle2 size={24} color="var(--success-text)" />
                <div>
                  <div style={{ fontWeight: 600, fontSize: 14, color: 'var(--success-text)' }}>Signed in</div>
                  <div style={{ color: 'var(--success-text)', fontSize: 13, fontFamily: 'JetBrains Mono, monospace' }}>
                    {savedEmail}
                  </div>
                </div>
              </div>
              <button
                className="btn btn-ghost btn-sm"
                onClick={handleSignOut}
                disabled={signingOut}
                style={{ whiteSpace: 'nowrap' }}
              >
                {signingOut ? '…' : <><LogOut size={16} /> Sign Out</>}
              </button>
            </div>
          ) : (
            <>
              {/* Requirements notice */}
              <div style={{ background: '#fffbeb', border: '1px solid #fcd34d', borderRadius: 8, padding: '12px 16px', marginBottom: 16, fontSize: 13 }}>
                <div style={{ fontWeight: 700, color: '#92400e', marginBottom: 6 }}>⚠️ Gmail App Password Requirements</div>
                <ul style={{ margin: 0, paddingLeft: 18, color: '#78350f', lineHeight: 1.8 }}>
                  <li>Must be a <strong>personal @gmail.com account</strong> (not a school or Workspace account)</li>
                  <li><strong>2-Step Verification</strong> must be enabled on your Google Account</li>
                  <li>App Passwords are created at <a href="https://myaccount.google.com/apppasswords" target="_blank" rel="noreferrer" style={{ color: '#0284c7', fontWeight: 600 }}>myaccount.google.com/apppasswords ↗</a></li>
                </ul>
                <div style={{ marginTop: 8, display: 'flex', gap: 8, flexWrap: 'wrap' }}>
                  <a href="https://myaccount.google.com/signinoptions/two-step-verification" target="_blank" rel="noreferrer"
                    style={{ fontSize: 12, background: '#fff', border: '1px solid #fcd34d', borderRadius: 6, padding: '4px 12px', color: '#92400e', textDecoration: 'none', fontWeight: 600 }}>
                    Enable 2-Step Verification ↗
                  </a>
                  <a href="https://myaccount.google.com/apppasswords" target="_blank" rel="noreferrer"
                    style={{ fontSize: 12, background: '#fff', border: '1px solid #fcd34d', borderRadius: 6, padding: '4px 12px', color: '#92400e', textDecoration: 'none', fontWeight: 600 }}>
                    Create App Password ↗
                  </a>
                </div>
              </div>
              <div className="form-row">
                <div className="form-group">
                  <label>Gmail address <span style={{ fontWeight: 400, color: 'var(--muted)', fontSize: 11 }}>(@gmail.com only)</span></label>
                  <input
                    type="email" placeholder="you@gmail.com"
                    value={senderEmail} onChange={e => setSenderEmail(e.target.value)}
                  />
                </div>
                <div className="form-group">
                  <label>App Password <span style={{ fontWeight: 400, color: 'var(--muted)' }}>(16 chars from App Passwords page)</span></label>
                  <input
                    type="password" placeholder="xxxx xxxx xxxx xxxx"
                    value={appPassword} onChange={e => setAppPassword(e.target.value)}
                    autoComplete="new-password"
                    onKeyDown={e => e.key === 'Enter' && authReady && !authSaving && handleSaveAuth()}
                  />
                </div>
              </div>
              <button
                className="btn btn-primary"
                onClick={handleSaveAuth}
                disabled={!authReady || authSaving}
                style={{ marginTop: 12 }}
              >
                {authSaving ? '…Saving' : '🔑 Save & Sign In'}
              </button>
              {authError && <div className="alert alert-danger mt-2">{authError}</div>}
            </>
          )}
        </div>
      </div>

      {/* ── Google Drive Integration Card ── */}
      <div className="card" style={{ marginTop: 24, borderColor: oauthDriveAvailable ? '#10b981' : driveAvailable ? '#f59e0b' : '#94a3b8' }}>
        <div className="card-header" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
          <span style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
            📁 Google Drive Integration (Automated Certificate Uploads)
          </span>
          {oauthDriveAvailable ? (
            <span style={{ fontSize: 12, backgroundColor: '#dcfce7', color: '#15803d', padding: '3px 10px', borderRadius: 12, fontWeight: 600 }}>✓ ACTIVE (OAuth2)</span>
          ) : driveAvailable ? (
            <span style={{ fontSize: 12, backgroundColor: '#fef3c7', color: '#92400e', padding: '3px 10px', borderRadius: 12, fontWeight: 600 }}>⚠ ACTIVE (Service Account)</span>
          ) : (
            <span style={{ fontSize: 12, backgroundColor: '#f1f5f9', color: '#64748b', padding: '3px 10px', borderRadius: 12, fontWeight: 500 }}>OPTIONAL</span>
          )}
        </div>
        <div className="card-body">

          {/* ── OPTION A: OAuth2 (Recommended) ── */}
          <div style={{ marginBottom: 20 }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 10 }}>
              <span style={{ fontWeight: 700, fontSize: 14, color: '#0f172a' }}>🔐 Method 1: Sign in with your Google Account</span>
              <span style={{ fontSize: 11, backgroundColor: '#dbeafe', color: '#1d4ed8', padding: '2px 8px', borderRadius: 10, fontWeight: 600 }}>RECOMMENDED</span>
            </div>
            <div style={{ fontSize: 13, color: '#64748b', marginBottom: 12, lineHeight: 1.5 }}>
              Files are uploaded owned by <strong>your</strong> Google account — uses your 15 GB storage quota. No storage limits!
            </div>

            {oauthDriveAvailable ? (
              <div style={{ backgroundColor: '#f0fdf4', border: '1px solid #86efac', borderRadius: 8, padding: '12px 16px' }}>
                <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', flexWrap: 'wrap', gap: 10 }}>
                  <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
                    <CheckCircle2 size={20} color="#16a34a" />
                    <div>
                      <div style={{ fontWeight: 600, fontSize: 13, color: '#15803d' }}>Connected as {oauthUserEmail || 'your Google Account'}</div>
                      <div style={{ fontSize: 12, color: '#4ade80' }}>Certificates upload directly to your Google Drive storage.</div>
                    </div>
                  </div>
                  <button className="btn btn-ghost btn-sm" onClick={handleRevokeOAuth} style={{ color: '#ef4444', fontSize: 12 }}>
                    Sign Out
                  </button>
                </div>
              </div>
            ) : (
              <div>
    {!oauthClientAvailable ? (
                  <div>
                    {/* Step-by-step guide for first-time setup */}
                    <div style={{ background: '#f8fafc', border: '1px solid #e2e8f0', borderRadius: 8, padding: '14px 16px', marginBottom: 12, fontSize: 13, lineHeight: 1.7 }}>
                      <div style={{ fontWeight: 700, color: '#0f172a', marginBottom: 8 }}>📋 One-time setup (3 steps):</div>
                      <ol style={{ margin: 0, paddingLeft: 20, color: '#334155' }}>
                        <li>Go to <a href="https://console.cloud.google.com/apis/credentials" target="_blank" rel="noreferrer" style={{ color: '#0284c7', fontWeight: 600 }}>Google Cloud Console → Credentials ↗</a></li>
                        <li>Click <strong>+ Create Credentials → OAuth 2.0 Client ID → Web application</strong></li>
                        <li>Under <strong>Authorized redirect URIs</strong>, add exactly:<br />
                          <code style={{ background: '#e0f2fe', color: '#0369a1', padding: '2px 8px', borderRadius: 4, fontSize: 12, userSelect: 'all', display: 'inline-block', marginTop: 4 }}>
                            {getBackendUrl()}/config/drive-auth/oauth-callback
                          </code>
                        </li>
                        <li>Download the <code>client_secret_...json</code> file and upload it below</li>
                      </ol>
                    </div>
                    <label className="btn btn-secondary" style={{ cursor: 'pointer', margin: 0, fontSize: 13 }}>
                      {oauthUploading ? 'Uploading…' : '📄 Upload client_secrets.json'}
                      <input type="file" accept=".json" style={{ display: 'none' }} onChange={handleOAuthClientUpload} disabled={oauthUploading} />
                    </label>
                  </div>
                ) : (
                  <div style={{ display: 'flex', alignItems: 'center', gap: 12, flexWrap: 'wrap' }}>
                    <span style={{ fontSize: 13, color: '#16a34a' }}>✓ client_secrets.json uploaded</span>
                    <button className="btn btn-primary" onClick={handleStartOAuth} style={{ fontSize: 13 }}>
                      🔗 Authorize with Google
                    </button>
                    <label className="btn btn-ghost btn-sm" style={{ cursor: 'pointer', margin: 0, fontSize: 12 }}>
                      Re-upload
                      <input type="file" accept=".json" style={{ display: 'none' }} onChange={handleOAuthClientUpload} />
                    </label>
                  </div>
                )}
                {oauthFlowStatus === 'in_progress' && (
                  <div style={{ marginTop: 10, background: '#eff6ff', border: '1px solid #bfdbfe', borderRadius: 8, padding: '12px 16px', fontSize: 13, color: '#1e40af' }}>
                    ⏳ <strong>Waiting for you to authorize in the new tab…</strong> Sign in with Google and grant Drive access, then return here.
                  </div>
                )}
                {oauthFlowStatus === 'error' && (
                  <div className="alert alert-danger" style={{ marginTop: 10, fontSize: 13 }}>❌ {oauthFlowError}</div>
                )}
              </div>
            )}
          </div>

          <div style={{ borderTop: '1px solid #e2e8f0', marginBottom: 16 }} />

          {/* ── OPTION B: Service Account (Advanced / Legacy) ── */}
          <div>
            <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 8 }}>
              <span style={{ fontWeight: 700, fontSize: 14, color: '#475569' }}>🤖 Method 2: Service Account Key</span>
              <span style={{ fontSize: 11, backgroundColor: '#f1f5f9', color: '#64748b', padding: '2px 8px', borderRadius: 10, fontWeight: 600 }}>ADVANCED</span>
            </div>

            {driveCredsExist ? (
              <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', flexWrap: 'wrap', gap: 10 }}>
                <div style={{ fontSize: 13, color: '#475569' }}>
                  {serviceAccountEmail ? (
                    <span>Key: <code style={{ background: '#e2e8f0', padding: '1px 4px', borderRadius: 3, fontSize: 12 }}>{serviceAccountEmail}</code></span>
                  ) : 'service_account.json loaded'}
                </div>
                <button className="btn btn-ghost btn-sm" onClick={handleDeleteDrive} disabled={driveUploading} style={{ color: '#ef4444', fontSize: 12 }}>
                  Remove Key
                </button>
              </div>
            ) : (
              <div>
                <div style={{ fontSize: 13, color: '#64748b', marginBottom: 8 }}>
                  ⚠️ Service accounts have <strong>0 storage quota</strong>. Only use if you have a Google Workspace Shared Drive.{' '}
                  <button type="button" onClick={() => setShowDriveGuide(!showDriveGuide)}
                    style={{ background: 'none', border: 'none', color: '#0284c7', fontWeight: 600, cursor: 'pointer', padding: 0, fontSize: 13 }}>
                    {showDriveGuide ? 'Hide guide ▲' : 'Setup guide ▼'}
                  </button>
                </div>
                {showDriveGuide && (
                  <div style={{ backgroundColor: '#f8fafc', border: '1px solid #e2e8f0', borderRadius: 6, padding: '12px 14px', marginBottom: 10, fontSize: 12, lineHeight: 1.6 }}>
                    <ol style={{ margin: 0, paddingLeft: 18 }}>
                      <li>Go to <a href="https://console.cloud.google.com/" target="_blank" rel="noreferrer" style={{ color: '#0284c7' }}>Google Cloud Console ↗</a></li>
                      <li>Enable <a href="https://console.cloud.google.com/apis/library/drive.googleapis.com" target="_blank" rel="noreferrer" style={{ color: '#0284c7' }}>Google Drive API ↗</a></li>
                      <li>Create a Service Account → Keys → Create new key (JSON)</li>
                      <li>Upload the downloaded .json file below</li>
                    </ol>
                  </div>
                )}
                <label className="btn btn-secondary btn-sm" style={{ cursor: 'pointer', margin: 0, fontSize: 13 }}>
                  {driveUploading ? 'Uploading…' : '📄 Upload service_account.json'}
                  <input type="file" accept=".json" style={{ display: 'none' }} onChange={handleDriveFileUpload} disabled={driveUploading} />
                </label>
              </div>
            )}
          </div>

          {/* Optional root folder override */}
          {driveAvailable && (
            <div style={{ borderTop: '1px solid #e2e8f0', paddingTop: 14, marginTop: 16 }}>
              <label style={{ fontWeight: 600, fontSize: 13, color: '#334155', display: 'block', marginBottom: 4 }}>
                Target Drive Folder URL / ID <span style={{ fontWeight: 400, color: '#64748b' }}>(optional — leave blank to auto-create)</span>
              </label>
              <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
                <input
                  type="text"
                  placeholder="https://drive.google.com/drive/folders/..."
                  value={driveFolderInput}
                  onChange={e => setDriveFolderInput(e.target.value)}
                  style={{ flex: 1, padding: '8px 12px', fontSize: 13, borderRadius: 6, border: '1px solid #cbd5e1' }}
                />
                <button className="btn btn-primary btn-sm" onClick={handleSaveDriveFolder} disabled={saveFolderLoading}>
                  {saveFolderLoading ? 'Saving…' : driveFolderSaved ? '✓ Saved' : 'Save'}
                </button>
              </div>
            </div>
          )}

          {driveError && <div className="alert alert-danger mt-2" style={{ marginTop: 10 }}>{driveError}</div>}
        </div>
      </div>
    </div>
  )
}

