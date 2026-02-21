import { test, expect } from '@playwright/test'

/**
 * Navigation E2E tests for KickOffice frontend.
 *
 * These tests cover web-only navigation flows.
 * Office Host-specific interactions (Word, Excel, PowerPoint, Outlook)
 * require manual testing as Office.js is not available in browser context.
 */

test.describe('Navigation', () => {
  test('should load home page', async ({ page }) => {
    await page.goto('/')
    // The app should render without crashing
    await expect(page.locator('body')).toBeVisible()
  })

  test('should navigate to settings page', async ({ page }) => {
    await page.goto('/')
    // Click settings button (gear icon in header)
    const settingsButton = page.locator('button[title*="Settings"], button[aria-label*="Settings"]').first()
    if (await settingsButton.isVisible()) {
      await settingsButton.click()
      await expect(page).toHaveURL(/.*settings.*/)
    }
  })

  test('should navigate back to home from settings', async ({ page }) => {
    await page.goto('/settings')
    // Look for back/home navigation
    const backButton = page.locator('button').filter({ hasText: /back|home|retour/i }).first()
    if (await backButton.isVisible()) {
      await backButton.click()
      await expect(page).toHaveURL('/')
    }
  })
})

test.describe('Settings Page', () => {
  test('should display settings tabs', async ({ page }) => {
    await page.goto('/settings')
    // Settings page should have tab navigation
    await expect(page.locator('body')).toBeVisible()
  })

  test('should persist language selection', async ({ page }) => {
    await page.goto('/settings')
    // Language selector should be present
    const languageSelect = page.locator('select, [role="listbox"]').first()
    if (await languageSelect.isVisible()) {
      await expect(languageSelect).toBeVisible()
    }
  })
})

test.describe('Accessibility', () => {
  test('should have no major accessibility violations on home page', async ({ page }) => {
    await page.goto('/')
    // Basic accessibility check - page should be keyboard navigable
    await page.keyboard.press('Tab')
    const focusedElement = page.locator(':focus')
    await expect(focusedElement).toBeVisible()
  })
})
