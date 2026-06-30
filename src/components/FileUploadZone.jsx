export default function FileUploadZone({ label, file, onFile }) {

    function handleDrop(e) {
      e.preventDefault()
      const dropped = e.dataTransfer.files[0]
      if (dropped) onFile(dropped)
    }
  
    return (
      <label
        onDragOver={(e) => e.preventDefault()}
        onDrop={handleDrop}
        className={`
          flex items-center gap-4 px-5 py-4 rounded-xl border cursor-pointer transition-all
          ${file
            ? 'border-green-500/30 bg-green-500/5'
            : 'border-[#DCD3F0] bg-[#F5F2FC] hover:border-[#8B7FC7]'
          }
        `}
      >
        <div className={`w-8 h-8 rounded-lg flex items-center justify-center
          ${file ? 'bg-green-500/20' : 'bg-[#F5F2FC]'}`}
        >
          {file ? '✓' : '↑'}
        </div>
  
        <div className="flex-1">
          <p className="text-sm font-medium text-[#4B3F72]">{label}</p>
          <p className="text-xs text-[#A99BC9] mt-0.5">
            {file ? file.name : 'drag and drop or click to upload'}
          </p>
        </div>
  
        {file && (
          <button
            type="button"
            onClick={(e) => {
              e.preventDefault()
              onFile(null)
            }}
            className="text-[#C7BEDF] hover:text-[#5B4B8A] transition"
          >
            ✕
          </button>
        )}
  
        <input
          type="file"
          accept=".xlsx"
          className="hidden"
          onChange={(e) => onFile(e.target.files[0])}
        />
      </label>
    )
  }