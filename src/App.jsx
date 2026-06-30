import { BrowserRouter, Routes, Route, NavLink } from 'react-router-dom'
import { LayoutDashboard, CheckSquare, History } from 'lucide-react'
import RunPage     from './pages/RunPage.jsx'
import ChecksPage  from './pages/ChecksPage.jsx'
import HistoryPage from './pages/HistoryPage.jsx'

const NAV = [
  { to: '/',        icon: LayoutDashboard, label: 'Run Plan'  },
  { to: '/checks',  icon: CheckSquare,     label: 'Checks'    },
  { to: '/history', icon: History,         label: 'History'   },
]

export default function App() {
  return (
    <BrowserRouter>
      <div className="min-h-screen bg-white text-[#4B3F72] flex">

        {/* Sidebar */}
        <aside className="w-56 shrink-0 bg-[#FBFAFE] border-r border-white/5 flex flex-col">

          {/* Logo area */}
          <div className="px-6 py-7 border-b border-white/5">
            <p className="text-xs font-semibold tracking-widest text-[#4B3F72]/30 uppercase mb-1">
              Supply Chain
            </p>
            <h1 className="text-lg font-bold text-[#4B3F72] leading-tight">
              Planning<br/>System
            </h1>
          </div>

          {/* Nav links */}
          <nav className="flex-1 px-3 py-4 space-y-1">
            {NAV.map(({ to, icon: Icon, label }) => (
              <NavLink
                key={to}
                to={to}
                end={to === '/'}
                className={({ isActive }) =>
                  `flex items-center gap-3 px-3 py-2.5 rounded-lg text-sm font-medium transition-all
                   ${isActive
                     ? 'bg-[#EFEAFA]/20 text-[#5B4B8A] border border-blue-500/20'
                     : 'text-[#4B3F72]/50 hover:text-[#4B3F72]/80 hover:border-[#EFEAFA]'
                   }`
                }
              >
                <Icon size={16} />
                {label}
              </NavLink>
            ))}
          </nav>

          <div className="px-6 py-4 border-t border-white/5">
            <p className="text-xs text-[#4B3F72]/20">v1.0</p>
          </div>
        </aside>

        {/* Page content */}
        <main className="flex-1 overflow-auto">
          <Routes>
            <Route path="/"        element={<RunPage />}     />
            <Route path="/checks"  element={<ChecksPage />}  />
            <Route path="/history" element={<HistoryPage />} />
          </Routes>
        </main>

      </div>
    </BrowserRouter>
  )
}