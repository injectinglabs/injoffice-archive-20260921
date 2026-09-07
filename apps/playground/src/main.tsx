import { StrictMode } from 'react'
import { createRoot } from 'react-dom/client'
import App from './App'
import './studio.css'
import './workbench.css'
import './design-system/tokens.css'
import './design-system/components.css'

createRoot(document.getElementById('root')!).render(
  <StrictMode>
    <App />
  </StrictMode>,
)
