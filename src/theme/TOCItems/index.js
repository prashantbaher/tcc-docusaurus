import React from 'react';
import TOCItems from '@theme-original/TOCItems';
import AdComponent from '@site/src/components/AdsenseVertical';

export default function TOCItemsWrapper(props) {
  return (
    <>
      <div>
        <TOCItems {...props} />
        <AdComponent />
      </div>
    </>
  );
}
