import type { AppearanceSurfaceDescriptor } from '@/infrastructure/appearance';

export const officeViewerAppearanceDescriptor: AppearanceSurfaceDescriptor = {
  id: 'office-viewer',
  parts: [
    { id: 'root' },
    { id: 'toolbar' },
    { id: 'info' },
    { id: 'container' },
    { id: 'content' },
    { id: 'loading' },
    { id: 'error' },
    { id: 'unsupported' },
  ],
};
