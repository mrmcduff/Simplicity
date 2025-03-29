/// <reference types="vite/client" />

/**
 * Define global Office and Excel types
 */
interface Window {
  Office: typeof Office;
  Excel: typeof Excel;
}
