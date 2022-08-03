import React from 'react';
import ComponentCreator from '@docusaurus/ComponentCreator';

export default [
  {
    path: '/__docusaurus/debug',
    component: ComponentCreator('/__docusaurus/debug', '3ab'),
    exact: true
  },
  {
    path: '/__docusaurus/debug/config',
    component: ComponentCreator('/__docusaurus/debug/config', 'd5a'),
    exact: true
  },
  {
    path: '/__docusaurus/debug/content',
    component: ComponentCreator('/__docusaurus/debug/content', 'bd0'),
    exact: true
  },
  {
    path: '/__docusaurus/debug/globalData',
    component: ComponentCreator('/__docusaurus/debug/globalData', '60e'),
    exact: true
  },
  {
    path: '/__docusaurus/debug/metadata',
    component: ComponentCreator('/__docusaurus/debug/metadata', '7a7'),
    exact: true
  },
  {
    path: '/__docusaurus/debug/registry',
    component: ComponentCreator('/__docusaurus/debug/registry', 'e53'),
    exact: true
  },
  {
    path: '/__docusaurus/debug/routes',
    component: ComponentCreator('/__docusaurus/debug/routes', 'b1c'),
    exact: true
  },
  {
    path: '/aboutme',
    component: ComponentCreator('/aboutme', '07a'),
    exact: true
  },
  {
    path: '/markdown-page',
    component: ComponentCreator('/markdown-page', '8ce'),
    exact: true
  },
  {
    path: '/resources',
    component: ComponentCreator('/resources', '27a'),
    exact: true
  },
  {
    path: '/search',
    component: ComponentCreator('/search', '2c5'),
    exact: true
  },
  {
    path: '/solidworks-cpp/tags',
    component: ComponentCreator('/solidworks-cpp/tags', 'af6'),
    exact: true
  },
  {
    path: '/solidworks-cpp/tags/managed-c',
    component: ComponentCreator('/solidworks-cpp/tags/managed-c', 'b13'),
    exact: true
  },
  {
    path: '/solidworks-cpp/tags/solidworks-c-api',
    component: ComponentCreator('/solidworks-cpp/tags/solidworks-c-api', 'd1c'),
    exact: true
  },
  {
    path: '/solidworks-cpp/tags/vc',
    component: ComponentCreator('/solidworks-cpp/tags/vc', '64c'),
    exact: true
  },
  {
    path: '/vba/tags',
    component: ComponentCreator('/vba/tags', 'ea3'),
    exact: true
  },
  {
    path: '/vba/tags/vba',
    component: ComponentCreator('/vba/tags/vba', '948'),
    exact: true
  },
  {
    path: '/vba/tags/vba-macro-testing',
    component: ComponentCreator('/vba/tags/vba-macro-testing', 'b67'),
    exact: true
  },
  {
    path: '/docs',
    component: ComponentCreator('/docs', '7bb'),
    routes: [
      {
        path: '/docs/category/tutorial---basics',
        component: ComponentCreator('/docs/category/tutorial---basics', 'd44'),
        exact: true,
        sidebar: "tutorialSidebar"
      },
      {
        path: '/docs/category/tutorial---extras',
        component: ComponentCreator('/docs/category/tutorial---extras', 'f09'),
        exact: true,
        sidebar: "tutorialSidebar"
      },
      {
        path: '/docs/intro',
        component: ComponentCreator('/docs/intro', 'd4f'),
        exact: true,
        sidebar: "tutorialSidebar"
      },
      {
        path: '/docs/tutorial-basics/congratulations',
        component: ComponentCreator('/docs/tutorial-basics/congratulations', '888'),
        exact: true,
        sidebar: "tutorialSidebar"
      },
      {
        path: '/docs/tutorial-basics/create-a-blog-post',
        component: ComponentCreator('/docs/tutorial-basics/create-a-blog-post', '94a'),
        exact: true,
        sidebar: "tutorialSidebar"
      },
      {
        path: '/docs/tutorial-basics/create-a-document',
        component: ComponentCreator('/docs/tutorial-basics/create-a-document', '4ca'),
        exact: true,
        sidebar: "tutorialSidebar"
      },
      {
        path: '/docs/tutorial-basics/create-a-page',
        component: ComponentCreator('/docs/tutorial-basics/create-a-page', '0db'),
        exact: true,
        sidebar: "tutorialSidebar"
      },
      {
        path: '/docs/tutorial-basics/deploy-your-site',
        component: ComponentCreator('/docs/tutorial-basics/deploy-your-site', 'd49'),
        exact: true,
        sidebar: "tutorialSidebar"
      },
      {
        path: '/docs/tutorial-basics/markdown-features',
        component: ComponentCreator('/docs/tutorial-basics/markdown-features', 'b3c'),
        exact: true,
        sidebar: "tutorialSidebar"
      },
      {
        path: '/docs/tutorial-extras/manage-docs-versions',
        component: ComponentCreator('/docs/tutorial-extras/manage-docs-versions', '950'),
        exact: true,
        sidebar: "tutorialSidebar"
      },
      {
        path: '/docs/tutorial-extras/translate-your-site',
        component: ComponentCreator('/docs/tutorial-extras/translate-your-site', '89a'),
        exact: true,
        sidebar: "tutorialSidebar"
      }
    ]
  },
  {
    path: '/solidworks-cpp',
    component: ComponentCreator('/solidworks-cpp', '8bc'),
    routes: [
      {
        path: '/solidworks-cpp/browse-and-open-document',
        component: ComponentCreator('/solidworks-cpp/browse-and-open-document', '2b9'),
        exact: true,
        sidebar: "tutorialSidebar"
      },
      {
        path: '/solidworks-cpp/cpp-prerequisite',
        component: ComponentCreator('/solidworks-cpp/cpp-prerequisite', '692'),
        exact: true,
        sidebar: "tutorialSidebar"
      },
      {
        path: '/solidworks-cpp/open-part-document',
        component: ComponentCreator('/solidworks-cpp/open-part-document', 'cef'),
        exact: true,
        sidebar: "tutorialSidebar"
      },
      {
        path: '/solidworks-cpp/open-solidworks',
        component: ComponentCreator('/solidworks-cpp/open-solidworks', 'd62'),
        exact: true,
        sidebar: "tutorialSidebar"
      },
      {
        path: '/solidworks-cpp/solidworks-Cpp-Api',
        component: ComponentCreator('/solidworks-cpp/solidworks-Cpp-Api', '15d'),
        exact: true,
        sidebar: "tutorialSidebar"
      }
    ]
  },
  {
    path: '/vba',
    component: ComponentCreator('/vba', 'b51'),
    routes: [
      {
        path: '/vba/browse-solidworks-file',
        component: ComponentCreator('/vba/browse-solidworks-file', '9e2'),
        exact: true
      },
      {
        path: '/vba/open-assembly-and-drawing-from-userform',
        component: ComponentCreator('/vba/open-assembly-and-drawing-from-userform', '8e2'),
        exact: true
      },
      {
        path: '/vba/open-part-from-userform',
        component: ComponentCreator('/vba/open-part-from-userform', '7d4'),
        exact: true
      },
      {
        path: '/vba/testing-open-assembly-and-drawing-document-macro',
        component: ComponentCreator('/vba/testing-open-assembly-and-drawing-document-macro', 'a1c'),
        exact: true
      },
      {
        path: '/vba/vba-arrays',
        component: ComponentCreator('/vba/vba-arrays', '62e'),
        exact: true,
        sidebar: "tutorialSidebar"
      },
      {
        path: '/vba/vba-assignment-statement-and-operator',
        component: ComponentCreator('/vba/vba-assignment-statement-and-operator', '625'),
        exact: true,
        sidebar: "tutorialSidebar"
      },
      {
        path: '/vba/vba-bug-finding',
        component: ComponentCreator('/vba/vba-bug-finding', '812'),
        exact: true
      },
      {
        path: '/vba/vba-bug-reduction-tips',
        component: ComponentCreator('/vba/vba-bug-reduction-tips', '5ab'),
        exact: true
      },
      {
        path: '/vba/vba-constant',
        component: ComponentCreator('/vba/vba-constant', '100'),
        exact: true,
        sidebar: "tutorialSidebar"
      },
      {
        path: '/vba/vba-controlling-flow-making-desicions',
        component: ComponentCreator('/vba/vba-controlling-flow-making-desicions', '3fa'),
        exact: true,
        sidebar: "tutorialSidebar"
      },
      {
        path: '/vba/vba-debugger',
        component: ComponentCreator('/vba/vba-debugger', '671'),
        exact: true
      },
      {
        path: '/vba/vba-declaring-and-scoping-of-variables',
        component: ComponentCreator('/vba/vba-declaring-and-scoping-of-variables', '0ed'),
        exact: true,
        sidebar: "tutorialSidebar"
      },
      {
        path: '/vba/vba-dialog-boxes',
        component: ComponentCreator('/vba/vba-dialog-boxes', '705'),
        exact: true
      },
      {
        path: '/vba/vba-executing-procedures',
        component: ComponentCreator('/vba/vba-executing-procedures', 'b9b'),
        exact: true,
        sidebar: "tutorialSidebar"
      },
      {
        path: '/vba/vba-functions',
        component: ComponentCreator('/vba/vba-functions', '1b6'),
        exact: true,
        sidebar: "tutorialSidebar"
      },
      {
        path: '/vba/vba-if-then-structure-select-case',
        component: ComponentCreator('/vba/vba-if-then-structure-select-case', '4c5'),
        exact: true,
        sidebar: "tutorialSidebar"
      },
      {
        path: '/vba/vba-inputbox-function',
        component: ComponentCreator('/vba/vba-inputbox-function', 'a44'),
        exact: true
      },
      {
        path: '/vba/vba-introduction',
        component: ComponentCreator('/vba/vba-introduction', 'af9'),
        exact: true,
        sidebar: "tutorialSidebar"
      },
      {
        path: '/vba/vba-looping',
        component: ComponentCreator('/vba/vba-looping', 'ce9'),
        exact: true
      },
      {
        path: '/vba/vba-more-function',
        component: ComponentCreator('/vba/vba-more-function', '1c3'),
        exact: true,
        sidebar: "tutorialSidebar"
      },
      {
        path: '/vba/vba-msgBox-function',
        component: ComponentCreator('/vba/vba-msgBox-function', '5af'),
        exact: true
      },
      {
        path: '/vba/vba-other-dialog',
        component: ComponentCreator('/vba/vba-other-dialog', '9a7'),
        exact: true
      },
      {
        path: '/vba/vba-programming-concepts',
        component: ComponentCreator('/vba/vba-programming-concepts', '0b6'),
        exact: true,
        sidebar: "tutorialSidebar"
      },
      {
        path: '/vba/vba-publc-static-variable-life',
        component: ComponentCreator('/vba/vba-publc-static-variable-life', 'e6c'),
        exact: true,
        sidebar: "tutorialSidebar"
      },
      {
        path: '/vba/vba-string-basic',
        component: ComponentCreator('/vba/vba-string-basic', '37a'),
        exact: true,
        sidebar: "tutorialSidebar"
      },
      {
        path: '/vba/vba-sub-and-function-procedure',
        component: ComponentCreator('/vba/vba-sub-and-function-procedure', 'fd1'),
        exact: true,
        sidebar: "tutorialSidebar"
      },
      {
        path: '/vba/vba-userform',
        component: ComponentCreator('/vba/vba-userform', 'd3a'),
        exact: true
      },
      {
        path: '/vba/vba-variable-scope',
        component: ComponentCreator('/vba/vba-variable-scope', '292'),
        exact: true,
        sidebar: "tutorialSidebar"
      },
      {
        path: '/vba/vba-variables',
        component: ComponentCreator('/vba/vba-variables', 'ca1'),
        exact: true,
        sidebar: "tutorialSidebar"
      },
      {
        path: '/vba/vbe-editor',
        component: ComponentCreator('/vba/vbe-editor', '6d1'),
        exact: true,
        sidebar: "tutorialSidebar"
      },
      {
        path: '/vba/vbe-windows',
        component: ComponentCreator('/vba/vbe-windows', '6a6'),
        exact: true,
        sidebar: "tutorialSidebar"
      }
    ]
  },
  {
    path: '/',
    component: ComponentCreator('/', '7ca'),
    exact: true
  },
  {
    path: '*',
    component: ComponentCreator('*'),
  },
];
