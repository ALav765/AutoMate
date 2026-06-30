import { useEffect, useState } from 'react'
import { Download, CheckCircle2 } from 'lucide-react'

export default function HistoryPage() {
  const [history, setHistory] = useState([])

  useEffect(() => {
    fetch('/api/history')
      .then(r => r.json())
      .then(d => setHistory(d.history))
  }, [])

  if (!history.length) return (
    <div className="p-8">
      <h2 className="text-2xl font-bold text-[#4B3F72] mb-2">History</h2>
      <p className="text-sm text-[#4B3F72]/30">No completed runs yet.</p>
    </div>
  )

  return (
    <div className="p-8 max-w-3xl mx-auto">
      <h2 className="text-2xl font-bold text-[#4B3F72] mb-8">History</h2>
      <div className="space-y-3">
        {history.map(job => (
          <div key={job.job_id}
            className="flex items-center justify-between bg-[#FBFAFE]
                       border border-white/5 rounded-xl px-5 py-4">
            <div className="flex items-center gap-4">
              <CheckCircle2 size={15} className="text-emerald-400 shrink-0" />
              <div>
                <p className="text-sm font-medium text-[#4B3F72]">Job {job.job_id}</p>
                <p className="text-xs text-[#4B3F72]/30 mt-0.5">
                  {job.summary?.passed}/{job.summary?.total} checks
                  · {job.summary?.loops} loop(s)
                  {job.summary?.business_alerts > 0 && ` · ${job.summary.business_alerts} alert(s)`}
                </p>
              </div>
            </div>
            
            <a  href={`/api/download/${job.job_id}`}
              download
              className="flex items-center gap-1.5 text-xs text-[#4B3F72]/40 hover:text-[#4B3F72] transition"
            >
              <Download size={13} /> Download
            </a>
          </div>
        ))}
      </div>
    </div>
  )
}