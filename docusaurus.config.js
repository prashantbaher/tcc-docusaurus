// @ts-check
// Note: type annotations allow type checking and IDEs autocompletion

const lightCodeTheme = require('prism-react-renderer/themes/github');
const darkCodeTheme = require('prism-react-renderer/themes/dracula');

/** @type {import('@docusaurus/types').Config} */
const config = {
  title: 'The CAD Coder',
  tagline: 'Free CAD Customization Tutorials for Mechanical Engineers.',
  url: 'https://thecadcoder.com/',
  baseUrl: '/',
  onBrokenLinks: 'throw',
  onBrokenMarkdownLinks: 'warn',
  favicon: 'assets/logo.png',
  organizationName: 'prashantbaher', // Usually your GitHub org/user name.
  projectName: 'tcc-blog', // Usually your repo name.
  deploymentBranch: 'master',

  presets: [
    [
      'classic',
      /** @type {import('@docusaurus/preset-classic').Options} */
      ({
        docs: {
          path: 'docs/demo',
          id: 'default',
          routeBasePath: 'docs',
          sidebarPath: require.resolve('./sidebars.js'),
          sidebarCollapsible: true,
        },
        blog: false,
        theme: {
          customCss: require.resolve('./src/css/custom.css'),
        },
      }),
    ],
  ],

  plugins: [
    [
      '@docusaurus/plugin-content-docs',
      {
        path: 'docs/vba',
        routeBasePath: 'vba',
        id: 'vba',
        sidebarPath: require.resolve('./sidebars/sidebars-vba.js'),
        sidebarCollapsible: true,
      },
    ],
    [
      '@docusaurus/plugin-content-docs',
      {
        path: 'docs/solidworks-cpp',
        routeBasePath: 'solidworks-cpp',
        id: 'solidworks-cpp',
        sidebarPath: require.resolve('./sidebars/sidebars-solidworks-cpp.js'),
        sidebarCollapsible: true,
      },
    ],
    [
      '@docusaurus/plugin-content-docs',
      {
        path: 'docs/solidworks-csharp',
        routeBasePath: 'solidworks-csharp',
        id: 'solidworks-csharp',
        sidebarPath: require.resolve('./sidebars/sidebars-solidworks-csharp.js'),
        sidebarCollapsible: true,
      },
    ],
    [
      '@docusaurus/plugin-content-docs',
      {
        path: 'docs/solidworks-macros',
        routeBasePath: 'solidworks-macros',
        id: 'solidworks-macros',
        sidebarPath: require.resolve('./sidebars/sidebars-solidworks-vba.js'),
        sidebarCollapsible: true,
      },
    ],
  ],

  themeConfig:
    /** @type {import('@docusaurus/preset-classic').ThemeConfig} */
    ({
      navbar: {
        title: 'The CAD Coder',
        logo: {
          alt: 'My Site Logo',
          src: 'assets/logo.png',
        },
        items: [
          {
            
            label: 'VBA',
            position: 'left',
            to: 'vba/vba-introduction',
          },
          {
            label: 'Solidworks API Tutorials',
            position: 'left',
            type: 'dropdown',
            items: [
              {
                label: 'Solidworks API + VBA',
                to: 'solidworks-macros/vba-in-solidworks'
              },
              {
                label: 'Solidworks API + C#',
                to: 'solidworks-csharp/solidworks-CSharp-Api'
              },
              {
                label: 'Solidworks API + C++',
                to: 'solidworks-cpp/solidworks-Cpp-Api'
              }
            ]
          },
          {
            label: 'Resources',
            position: 'right',
            to: '/resources'
          },
          {
            label: 'About Me',
            position: 'right',
            to: '/aboutme'
          },
        ],
      },
      footer: {
        style: 'light',
        links: [
          {
            title: 'Docs',
            items: [
              {
                label: 'Tutorial',
                to: '/docs/intro',
              },
            ],
          },
          {
            title: 'Community',
            items: [
              {
                label: 'Stack Overflow',
                href: 'https://stackoverflow.com/questions/tagged/docusaurus',
              },
              {
                label: 'Discord',
                href: 'https://discordapp.com/invite/docusaurus',
              },
              {
                label: 'Twitter',
                href: 'https://twitter.com/docusaurus',
              },
            ],
          },
          {
            title: 'More',
            items: [
              {
                label: 'Blog',
                to: '/blog',
              },
              {
                label: 'GitHub',
                href: 'https://github.com/facebook/docusaurus',
              },
            ],
          },
        ],
        copyright: `Copyright © ${new Date().getFullYear()} The CAD Coder.`,
      },
      prism: {
        theme: require('prism-react-renderer/themes/vsLight'),
        darkTheme: require('prism-react-renderer/themes/vsDark'),
        additionalLanguages: ['cpp', 'clike' ,'csharp', 'visual-basic'],
        showLineNumbers: false
      },
      docs: {
        sidebar: {hideable: true}
      }
    }),
    themes: [
      [
        require.resolve("@easyops-cn/docusaurus-search-local"),
        {
          hashed: true,
        },
      ],
    ],
};

module.exports = config;
