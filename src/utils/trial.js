/*
 * Excel Data Cleaner Add-in
 * Utility functions (formerly trial management)
 */

// This file has been simplified to remove all trial/premium functionality
// It now only maintains compatibility with code that might call these functions

/**
 * Always returns true since we've removed trial limitations
 * @returns {boolean} Always true
 */
export function checkTrialLimits() {
  return true;
}

/**
 * Always returns 0 since we've removed the operation counter
 * @returns {number} Always 0
 */
export function getOperationCount() {
  return 0;
}

/**
 * No-op function that does nothing
 */
export function incrementOperationCount() {
  // No operation needed
}

/**
 * Always returns true since all features are now included
 * @returns {boolean} Always true
 */
export function isPremiumUser() {
  return true;
}

/**
 * Simplified function just to maintain compatibility
 */
export function handleUpgrade() {
  // No operation needed
}

/**
 * No-op function that does nothing
 */
export function resetTrialCounter() {
  // No operation needed
}

/**
 * No-op function that does nothing
 * @param {boolean} isPremium - Ignored parameter
 */
export function setPremiumStatus(isPremium) {
  // No operation needed
}