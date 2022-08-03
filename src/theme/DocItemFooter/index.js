import React from 'react';
import DocItemFooter from '@theme-original/DocItemFooter';
import AdComponent from '@site/src/components/Adsense';

export default function DocItemFooterWrapper(props) {
  return (
    <>
      <br />
      <script>window.location.reload(true);</script> 
      <AdComponent />
      <DocItemFooter {...props} />
    </>
  );
}
