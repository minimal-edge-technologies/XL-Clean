/*
 * Excel Data Cleaner Add-in
 * UI utilities
 */

/**
 * Show a message to the user
 * @param {string} message - The message to show
 * @param {string} type - Message type (info, success, error)
 */
export function showMessage(message, type = "info") {
  console.log(`[${type}] ${message}`);
  
  // Create notification container if it doesn't exist
  let container = document.getElementById("notification-container");
  if (!container) {
    container = document.createElement("div");
    container.id = "notification-container";
    container.className = "notification-container";
    document.body.appendChild(container);
  }
  
  // Create notification element
  const notification = document.createElement("div");
  notification.className = `toast-notification ${type}`;
  
  // Add icon based on type
  const icon = document.createElement("i");
  icon.className = `fas fa-${type === 'success' ? 'check-circle' : type === 'error' ? 'exclamation-circle' : 'info-circle'}`;
  notification.appendChild(icon);
  
  // Add content
  const content = document.createElement("div");
  content.className = "toast-content";
  
  const title = document.createElement("div");
  title.className = "toast-title";
  title.textContent = type.charAt(0).toUpperCase() + type.slice(1);
  content.appendChild(title);
  
  const messageEl = document.createElement("div");
  messageEl.className = "toast-message";
  messageEl.textContent = message;
  content.appendChild(messageEl);
  
  notification.appendChild(content);
  
  // Add close button
  const closeBtn = document.createElement("button");
  closeBtn.className = "toast-close";
  closeBtn.innerHTML = '<i class="fas fa-times"></i>';
  closeBtn.addEventListener("click", () => {
    notification.remove();
  });
  notification.appendChild(closeBtn);
  
  // Add to container
  container.appendChild(notification);
  
  // Remove after 5 seconds
  setTimeout(() => {
    if (notification.parentNode === container) {
      notification.style.opacity = "0";
      notification.style.transform = "translateX(100%)";
      notification.style.transition = "opacity 0.3s ease, transform 0.3s ease";
      
      setTimeout(() => {
        if (notification.parentNode === container) {
          container.removeChild(notification);
        }
      }, 300);
    }
  }, 5000);
}

/**
 * Show error message
 * @param {Error} error - The error object
 */
export function showError(error) {
  console.error(error);
  showMessage(`Error: ${error.message}`, "error");
}

// Removed premium dialog functions since they're no longer needed

/**
 * Show the loading overlay
 * @param {string} message - Message to display while loading
 */
export function showLoading(message = "Processing data...") {
  const loadingOverlay = document.getElementById("loading-overlay");
  const loadingText = loadingOverlay.querySelector(".loading-text");
  
  // Set custom message if provided
  if (message) {
    loadingText.textContent = message;
  }
  
  // Show the overlay
  loadingOverlay.style.display = "flex";
  loadingOverlay.style.opacity = "1";
  
  // Prevent scrolling while loading
  document.body.style.overflow = "hidden";
}

/**
 * Hide the loading overlay
 */
export function hideLoading() {
  const loadingOverlay = document.getElementById("loading-overlay");
  
  // Fade out the overlay
  loadingOverlay.style.opacity = "0";
  
  // Wait for the transition to complete before hiding completely
  setTimeout(() => {
    loadingOverlay.style.display = "none";
    document.body.style.overflow = "auto";
  }, 300);
}

// For backwards compatibility
export function showUpgradeDialog() {
  // No-op function
}

export function hideUpgradeDialog() {
  // No-op function
}