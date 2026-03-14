import type { ToolDefinition } from '@/types';
import { getWordToolDefinitions } from '@/utils/wordTools';
import { getExcelToolDefinitions } from '@/utils/excelTools';
import { getPowerPointToolDefinitions } from '@/utils/powerpointTools';
import { getOutlookToolDefinitions } from '@/utils/outlookTools';

/**
 * ARCH-M1: ToolProviderRegistry
 *
 * Provides a host-agnostic registry for tool definitions.
 * Eliminates hard-coded imports and switch logic in useAgentLoop.
 *
 * Benefits:
 * - Adding new Office hosts (OneNote, Teams, etc.) requires only registration, no loop changes
 * - Centralized tool provider management
 * - Easier testing (mock registry for unit tests)
 * - Clear separation of concerns
 */

type ToolProvider = () => ToolDefinition[];

interface ToolProviderRegistry {
  register(hostName: string, provider: ToolProvider): void;
  getTools(hostName: string): ToolDefinition[];
  getRegisteredHosts(): string[];
  hasProvider(hostName: string): boolean;
}

class ToolProviderRegistryImpl implements ToolProviderRegistry {
  private providers = new Map<string, ToolProvider>();

  /**
   * Register a tool provider for a specific host
   * @param hostName - Office host name (e.g., 'Word', 'Excel', 'PowerPoint', 'Outlook')
   * @param provider - Function that returns tool definitions for this host
   */
  register(hostName: string, provider: ToolProvider): void {
    const normalizedName = hostName.toLowerCase();
    if (this.providers.has(normalizedName)) {
      console.warn(`[ToolProviderRegistry] Overwriting existing provider for host: ${hostName}`);
    }
    this.providers.set(normalizedName, provider);
  }

  /**
   * Get tool definitions for a specific host
   * @param hostName - Office host name
   * @returns Array of tool definitions, or empty array if host not registered
   */
  getTools(hostName: string): ToolDefinition[] {
    const normalizedName = hostName.toLowerCase();
    const provider = this.providers.get(normalizedName);

    if (!provider) {
      console.warn(
        `[ToolProviderRegistry] No provider registered for host: ${hostName}. Available hosts: ${this.getRegisteredHosts().join(', ')}`,
      );
      return [];
    }

    return provider();
  }

  /**
   * Get list of all registered host names
   */
  getRegisteredHosts(): string[] {
    return Array.from(this.providers.keys());
  }

  /**
   * Check if a provider is registered for a host
   */
  hasProvider(hostName: string): boolean {
    return this.providers.has(hostName.toLowerCase());
  }
}

// Singleton instance
const registry = new ToolProviderRegistryImpl();

// Register all providers at module initialization using static ES imports
registry.register('Word', getWordToolDefinitions);
registry.register('Excel', getExcelToolDefinitions);
registry.register('PowerPoint', getPowerPointToolDefinitions);
registry.register('Outlook', getOutlookToolDefinitions);

/**
 * Get the global tool provider registry
 */
export function getToolProviderRegistry(): ToolProviderRegistry {
  return registry;
}

/**
 * Get tools for the current Office host based on host flags
 * @param hostFlags - Object with isOutlook, isPowerPoint, isExcel flags
 * @returns Tool definitions for the current host
 */
export function getToolsForHost(hostFlags: {
  isOutlook: boolean;
  isPowerPoint: boolean;
  isExcel: boolean;
}): ToolDefinition[] {
  const { isOutlook, isPowerPoint, isExcel } = hostFlags;

  let hostName = 'Word'; // default
  if (isOutlook) hostName = 'Outlook';
  else if (isPowerPoint) hostName = 'PowerPoint';
  else if (isExcel) hostName = 'Excel';

  return registry.getTools(hostName);
}
