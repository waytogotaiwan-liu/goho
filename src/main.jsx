import './storage.js'
import React from 'react'
import { createRoot } from 'react-dom/client'
import GohoSystem from './App.jsx'

createRoot(document.getElementById('root')).render(
  <React.StrictMode>
    <GohoSystem />
  </React.StrictMode>
)
