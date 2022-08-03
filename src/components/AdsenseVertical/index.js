import React from 'react';

export default class HomePage extends React.Component {
  componentDidMount() {
    const installGoogleAds = () => {
      const elem = document.createElement("script");
      elem.src =
        "//pagead2.googlesyndication.com/pagead/js/adsbygoogle.js";
      elem.async = true;
      elem.defer = true;
      document.body.insertBefore(elem, document.body.firstChild);
    };
    installGoogleAds();
    (window.adsbygoogle = window.adsbygoogle || []).push({});
  }

  render() {
    return (
      <ins className='adsbygoogle'
           style={{ display: 'block' }}
           data-ad-client='ca-pub-8158659264340002'
           data-ad-slot='4311368222'
           data-ad-format='auto'
           data-full-width-responsive="true" />
    );
  }
}