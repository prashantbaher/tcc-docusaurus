import React from 'react';
import clsx from 'clsx';
import Layout from '@theme/Layout';
import Link from '@docusaurus/Link';
import useDocusaurusContext from '@docusaurus/useDocusaurusContext';
import styles from './index.module.css';
import HomepageFeatures from '@site/src/components/HomepageFeatures';
import Head from '@docusaurus/Head';

function HomepageHeader() {
  const {siteConfig} = useDocusaurusContext();
  return (
    <header className={clsx('hero hero--primary', styles.heroBanner)}>
      <div className="container">
        <h1 className="hero__title">{siteConfig.title}</h1>
        <p className="hero__subtitle">{siteConfig.tagline}</p>
        <div className={styles.buttons}>
          <Link
            className="button button--secondary button--lg"
            to="/docs/intro">
            Get Started
          </Link>
        </div>
      </div>
    </header>
  );
}

export default function Home() {
  const {siteConfig} = useDocusaurusContext();
  return (
    <>
    <Head>
      <meta name="google-site-verification" content="OeSBbUv_kXDkDxaeoLixjE1BYxlbQRINt58H1UGbSMY" />
    </Head>
    <script async src="//pagead2.googlesyndication.com/pagead/js/adsbygoogle.js"></script>
     <script src="/assets/js/_main.js"></script>
    <Layout
      title={`${siteConfig.title}`}
      description="Free CAD Customization Tutorials for Mechanical Engineers.">
      <HomepageHeader />
      <main>
        <HomepageFeatures />
      </main>
    </Layout>
    </>
  );
}
