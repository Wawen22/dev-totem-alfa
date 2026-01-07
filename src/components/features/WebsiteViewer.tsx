import React from 'react';

interface WebsiteViewerProps {
  url: string;
}

export function WebsiteViewer({ url }: WebsiteViewerProps) {
  return (
    <div style={{ flex: 1, width: '100%', height: '100%', display: 'flex', flexDirection: 'column', backgroundColor: '#fff', borderRadius: '8px', overflow: 'hidden' }}>
       <iframe 
          src={url} 
          style={{ flex: 1, border: 'none', width: '100%', height: '100%', minHeight: 'calc(100vh - 200px)' }}
          title="Sito Web"
          sandbox="allow-same-origin allow-scripts allow-popups allow-forms"
        />
    </div>
  );
}
