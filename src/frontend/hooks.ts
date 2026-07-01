import { useEffect, useRef, useState } from 'react';

export type PluginComponentProps<
  TState = Record<string, unknown>,
  TConfig = Record<string, unknown>,
> = {
  pluginName: string;
  pluginState?: TState;
  pluginConfig?: TConfig;
  setPluginConfig?: (path: string, value: unknown) => Promise<void>;
  onAction: (action: string, data?: unknown) => void;
  onClose?: () => void;
  props?: Record<string, unknown>;
};

/**
 * Kai renders plugin panels inside a host that is itself the scroll container,
 * so `h-full` grows with content instead of constraining it. This hook measures
 * the available height (from the panel's top to the bottom of the nearest
 * scrollable ancestor) and returns a pixel value to pin the root to, so inner
 * `overflow-y-auto` columns can scroll independently.
 */
export function usePanelHeight(min = 320): [React.RefObject<HTMLDivElement | null>, number | null] {
  const ref = useRef<HTMLDivElement>(null);
  const [height, setHeight] = useState<number | null>(null);

  useEffect(() => {
    const panel = ref.current;
    if (!panel || typeof window === 'undefined') return;

    const findScrollParent = () => {
      let p = panel.parentElement;
      while (p) {
        const oy = window.getComputedStyle(p).overflowY;
        if ((oy === 'auto' || oy === 'scroll') && p.clientHeight > 0) return p;
        p = p.parentElement;
      }
      return null;
    };

    const scrollParent = findScrollParent();
    const measure = () => {
      const rect = panel.getBoundingClientRect();
      const parentRect = scrollParent?.getBoundingClientRect();
      const parentStyle = scrollParent ? window.getComputedStyle(scrollParent) : null;
      const parentBottom = parentRect?.bottom ?? window.innerHeight;
      const padBottom = parentStyle ? Number.parseFloat(parentStyle.paddingBottom) || 0 : 0;
      setHeight(Math.max(min, Math.floor(parentBottom - rect.top - padBottom)));
    };

    measure();
    window.addEventListener('resize', measure);
    const ro = typeof ResizeObserver !== 'undefined' ? new ResizeObserver(measure) : null;
    ro?.observe(panel);
    if (scrollParent) ro?.observe(scrollParent);

    return () => {
      window.removeEventListener('resize', measure);
      ro?.disconnect();
    };
  }, [min]);

  return [ref, height];
}
