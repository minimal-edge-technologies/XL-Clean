/*
 * Excel Data Cleaner Add-in
 * Accessibility helper functions
 */

/**
 * Initialize accessibility features
 */
export function initializeAccessibility() {
    // Add skip links for keyboard navigation
    addSkipLinks();
    
    // Ensure proper focus management
    setupFocusManagement();
    
    // Add ARIA attributes where needed
    enhanceARIASupport();
  }
  
  /**
   * Add skip links for keyboard navigation
   */
  function addSkipLinks() {
    // Skip links already added in HTML
    // This is just a placeholder for future enhancements
  }
  
  /**
   * Set up proper focus management
   */
  function setupFocusManagement() {
    // When dialogs open, save previous focus point
    const dialogs = document.querySelectorAll('.ms-Dialog');
    
    dialogs.forEach(dialog => {
      // Monitor for dialog visibility changes
      const observer = new MutationObserver((mutations) => {
        mutations.forEach((mutation) => {
          if (mutation.attributeName === 'style') {
            const isVisible = dialog.style.display === 'block';
            
            if (isVisible) {
              // When dialog opens, focus the first focusable element
              setTimeout(() => {
                const focusable = dialog.querySelector('button, [href], input, select, textarea, [tabindex]:not([tabindex="-1"])');
                if (focusable) {
                  focusable.focus();
                }
              }, 100);
            }
          }
        });
      });
      
      observer.observe(dialog, { attributes: true });
    });
  }
  
  /**
   * Enhance ARIA support for dynamic elements
   */
  function enhanceARIASupport() {
    // Add appropriate ARIA roles and states to UI elements
    
    // Ensure loading overlay has proper ARIA attributes
    const loadingOverlay = document.getElementById('loading-overlay');
    if (loadingOverlay) {
      loadingOverlay.setAttribute('role', 'alert');
      loadingOverlay.setAttribute('aria-live', 'assertive');
    }
    
    // Set up message container for screen readers
    const messageContainer = document.getElementById('message-container');
    if (messageContainer) {
      messageContainer.setAttribute('aria-live', 'polite');
    }
  }