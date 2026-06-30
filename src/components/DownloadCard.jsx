import { Download, FileSpreadsheet } from 'lucide-react'

export default function DownloadCard({ jobId }) {
  const url = '/api/download/' + jobId

  return (
    
    <a  href={url}
      download="Integrated_Supply_Plan.xlsx"
      className="flex items-center gap-4 bg-teal-500/10 hover:bg-teal-500/15 border border-teal-500/20 rounded-2xl px-6 py-5 transition group"
    >
      <div className="w-10 h-10 rounded-xl bg-teal-500/20 flex items-center justify-center">
        <FileSpreadsheet size={18} className="text-teal-400" />
      </div>
      <div className="flex-1">
        <p className="text-sm font-semibold text-white">Integrated_Supply_Plan.xlsx</p>
        <p className="text-xs text-white/40 mt-0.5">Ready · includes Alerts and Flags sheet</p>
      </div>
      <Download size={16} className="text-teal-400 group-hover:translate-y-0.5 transition-transform" />
    </a>
  )
}
