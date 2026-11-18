import { useLayoutEffect, useRef, useState } from 'react';

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

export function measurePanelHeight(panel: HTMLElement, scrollParent: HTMLElement | null): number {
  const rect = panel.getBoundingClientRect();
  if (!scrollParent) return Math.max(0, Math.floor(window.innerHeight - rect.top));

  // Use the unscrolled position; otherwise host scrolling makes the panel grow.
  const top = rect.top - scrollParent.getBoundingClientRect().top
    - scrollParent.clientTop + scrollParent.scrollTop;
  let bottomInset = Number.parseFloat(window.getComputedStyle(scrollParent).paddingBottom) || 0;
  for (let p = panel.parentElement; p && p !== scrollParent; p = p.parentElement) {
    const style = window.getComputedStyle(p);
    bottomInset += (Number.parseFloat(style.paddingBottom) || 0)
      + (Number.parseFloat(style.borderBottomWidth) || 0)
      + (Number.parseFloat(style.marginBottom) || 0);
  }
  return Math.max(0, Math.floor(scrollParent.clientHeight - top - bottomInset));
}

/** Constrain the whole plugin, including navigation, inside Kai's scrolling host. */
export function usePanelHeight(): [React.RefObject<HTMLDivElement | null>, number] {
  const ref = useRef<HTMLDivElement>(null);
  const [height, setHeight] = useState(0);

  useLayoutEffect(() => {
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
    const measure = () => setHeight(measurePanelHeight(panel, scrollParent));

    measure();
    window.addEventListener('resize', measure);
    const ro = typeof ResizeObserver !== 'undefined' ? new ResizeObserver(measure) : null;
    ro?.observe(panel);
    for (let p = panel.parentElement; p; p = p.parentElement) {
      ro?.observe(p);
      if (p === scrollParent) break;
    }

    return () => {
      window.removeEventListener('resize', measure);
      ro?.disconnect();
    };
  }, []);

  return [ref, height];
}
