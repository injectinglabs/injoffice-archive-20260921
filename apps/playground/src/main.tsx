import { StrictMode } from 'react'
import { createRoot } from 'react-dom/client'
import App from './App'
import './studio.css'
import './workbench.css'

createRoot(document.getElementById('root')!).render(
  <StrictMode>
    <App />
  </StrictMode>,
)
