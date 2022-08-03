import React from 'react';
import clsx from 'clsx';
import styles from './styles.module.css';

const FeatureList = [
  {
    title: 'Easy to Understand',
    Svg: require('@site/static/assets/blogger.svg').default,
    description: (
      <>
        This website prepared for beginners especially for <strong>Mechanical engineers</strong>. 
        Each article in this website is <strong>self-contained</strong>. These articles written 
        in <strong>easy to understand language</strong> so that normal people can understand it.
      </>
    ),
  },
  {
    title: 'Focus on What Matters',
    Svg: require('@site/static/assets/developer.svg').default,
    description: (
      <>
        Each article has <strong>Table of Content</strong> on right side. 
        By using this, you can directly go to your interested area of article. <br/>
        <em>This will help you to focus what matters to you.</em>
      </>
    ),
  },
  {
    title: 'Current Path',
    Svg: require('@site/static/assets/opensource.svg').default,
    description: (
      <>
        Currently, I am focusing completing SOLIDWORKS VBA API. <br/>
        After completing it I will start <strong>SOLIDWORKS C# API</strong> tutorial articles. 
        Let me know what you think about it.
      </>
    ),
  },
];

function Feature({Svg, title, description}) {
  return (
    <div className={clsx('col col--4')}>
      <div className="text--center">
        <Svg className={styles.featureSvg} role="img" />
      </div>
      <div className="text--center padding-horiz--md">
        <h3>{title}</h3>
        <hr />
        <p>{description}</p>
      </div>
    </div>
  );
}

export default function HomepageFeatures() {
  return (
    <section className={styles.features}>
      <div className="container">
        <div className="row">
          {FeatureList.map((props, idx) => (
            <Feature key={idx} {...props} />
          ))}
        </div>
      </div>
    </section>
  );
}
