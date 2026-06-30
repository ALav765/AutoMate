import { useState } from 'react'
import { Play, RefreshCw } from 'lucide-react'
import FileUploadZone  from '../components/FileUploadZone'
import RunStatus       from '../components/RunStatus'
import ChecklistPanel  from '../components/ChecklistPanel'
import DownloadCard    from '../components/DownloadCard'

// The 5 required files — defined outside the component
// so this array isn't recreated on every render
const REQUIRED_FILES = [
  { key: 'forecast',      label: 'Forecast & FG SOH'      },
  { key: 'bom',           label: 'Bill of Materials'       },
  { key: 'component_soh', label: 'Component SOH'           },
  { key: 'prices',        label: 'Component Prices'        },
  { key: 'vendor',        label: 'Vendor Master'           },
]

export default function RunPage() {
  const [files, setFiles]           = useState({})
  const [monthLabel, setMonthLabel] = useState('')
  const [jobId, setJobId]           = useState(null)
  const [status, setStatus]         = useState(null)
  const [events, setEvents]         = useState([])
  const [checkResults, setCheckResults] = useState(null)

  // All 5 files uploaded and month label filled in
  const ready = REQUIRED_FILES.every(f => files[f.key]) && monthLabel.trim()

  async function handleRun() {
    // Reset everything from a previous run
    setJobId(null)
    setEvents([])
    setCheckResults(null)
    setStatus('running')

    // Build the form data to send to the backend
    const form = new FormData()
    REQUIRED_FILES.forEach(({ key }) => form.append(key, files[key]))
    form.append('month_label', monthLabel)

    // POST to backend, get a job ID back immediately
    const res  = await fetch('/api/run', { method: 'POST', body: form })
    const { job_id } = await res.json()
    setJobId(job_id)

    // Open an SSE connection to stream live updates
    const es = new EventSource(`/api/status/${job_id}`)
    es.onmessage = (e) => {
      const event = JSON.parse(e.data)
      setEvents(prev => [...prev, event])
      if (event.type === 'check_results') setCheckResults(event.results)
      if (event.type === 'done')  { setStatus('done');  es.close() }
      if (event.type === 'error') { setStatus('error'); es.close() }
    }
  }

  function handleReset() {
    setFiles({})
    setMonthLabel('')
    setJobId(null)
    setStatus(null)
    setEvents([])
    setCheckResults(null)
  }

  return (
    <div className="p-8 max-w-3xl mx-auto">

      <div className="mb-8">
        <h2 className="text-2xl font-bold text-[#4B3F72]">Run Supply Plan</h2>
        <p className="text-sm text-[#4B3F72]/40 mt-1">
          Upload your five input files and run the full pipeline.
        </p>
      </div>

      {/* Upload form — only shown before running */}
      {!status && (
        <div className="space-y-6">

          {/* Month label input */}
          <div>
            <label className="block text-xs font-semibold text-[#4B3F72]/40 uppercase tracking-wider mb-2">
              Plan Month
            </label>
            <input
              type="text"
              placeholder="e.g. Oct-25"
              value={monthLabel}
              onChange={(e) => setMonthLabel(e.target.value)}
              className="w-44 border-[#EFEAFA] border border-[#DCD3F0] rounded-lg px-4 py-2.5
                         text-sm text-[#4B3F72] placeholder:text-[#4B3F72]/20
                         focus:outline-none focus:border-[#8B7FC7] transition"
            />
          </div>

          {/* File upload zones */}
          <div>
            <label className="block text-xs font-semibold text-[#4B3F72]/40 uppercase tracking-wider mb-3">
              Input Files
            </label>
            <div className="space-y-2">
              {REQUIRED_FILES.map(f => (
                <FileUploadZone
                  key={f.key}
                  label={f.label}
                  file={files[f.key]}
                  onFile={(file) => setFiles(prev => ({ ...prev, [f.key]: file }))}
                />
              ))}
            </div>
          </div>

          {/* Run button */}
          <button
            onClick={handleRun}
            disabled={!ready}
            className="flex items-center gap-2 bg-[#8B7FC7] hover:bg-[#7A6BB8]
                       disabled:bg-[#EFEAFA] disabled:text-[#C7BEDF]
                       text-[#4B3F72] font-semibold px-6 py-3 rounded-xl transition text-sm"
          >
            <Play size={15} />
            Run Pipeline
          </button>

        </div>
      )}

      {/* Live status + results — shown while running or after done */}
      {status && (
        <div className="space-y-6">
          <RunStatus events={events} status={status} />
          {checkResults && <ChecklistPanel results={checkResults} />}
          {status === 'done' && jobId && <DownloadCard jobId={jobId} />}
          <button
            onClick={handleReset}
            className="flex items-center gap-2 text-sm text-[#4B3F72]/30 hover:text-[#4B3F72]/60 transition"
          >
            <RefreshCw size={14} /> Start a new run
          </button>
        </div>
      )}

    </div>
  )
}