import { StrictMode } from 'react'
import { createRoot } from 'react-dom/client'
import './index.css'
import App from './App.tsx'

// Wait for DOM to be ready
const initApp = () => {
  const rootElement = document.getElementById('root');
  
  if (rootElement) {
    try {
      createRoot(rootElement).render(
        <StrictMode>
          <App />
        </StrictMode>,
      );
    } catch (error) {
      console.error('Error rendering app:', error);
      // Fallback: show basic error message
      rootElement.innerHTML = '<div style="padding: 20px; text-align: center;"><h1>Loading Error</h1><p>Please refresh the page.</p></div>';
    }
  } else {
    console.error('Root element not found!');
  }
};

// Ensure DOM is loaded
if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', initApp);
} else {
  initApp();
}
