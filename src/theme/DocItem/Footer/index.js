import React from 'react';
import Footer from '@theme-original/DocItem/Footer';
import AdComponent from '@site/src/components/Adsense';

export default function FooterWrapper(props) {
  return (
    <>
      <br />
      <script>window.location.reload(true);</script> 
      <AdComponent />
      <Footer {...props} />
    </>
  );
}
