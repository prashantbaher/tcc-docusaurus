/*
 * AUTOGENERATED - DON'T EDIT
 * Your edits in this file will be overwritten in the next build!
 * Modify the docusaurus.config.js file at your site's root instead.
 */
export default {
  "title": "The CAD Coder",
  "tagline": "Free CAD Customization Tutorials for Mechanical Engineers.",
  "url": "https://thecadcoder.com/",
  "baseUrl": "/",
  "onBrokenLinks": "throw",
  "onBrokenMarkdownLinks": "warn",
  "favicon": "assets/logo.png",
  "organizationName": "prashantbaher",
  "projectName": "tcc-blog",
  "deploymentBranch": "master",
  "presets": [
    [
      "classic",
      {
        "docs": {
          "path": "docs/demo",
          "id": "default",
          "routeBasePath": "docs",
          "sidebarPath": "D:\\website\\docs-website\\sidebars.js",
          "sidebarCollapsible": true
        },
        "blog": false,
        "theme": {
          "customCss": "D:\\website\\docs-website\\src\\css\\custom.css"
        }
      }
    ]
  ],
  "plugins": [
    [
      "@docusaurus/plugin-content-docs",
      {
        "path": "docs/vba",
        "routeBasePath": "vba",
        "id": "vba",
        "sidebarPath": "D:\\website\\docs-website\\sidebars\\sidebars-vba.js",
        "sidebarCollapsible": true
      }
    ],
    [
      "@docusaurus/plugin-content-docs",
      {
        "path": "docs/solidworks-cpp",
        "routeBasePath": "solidworks-cpp",
        "id": "solidworks-cpp",
        "sidebarPath": "D:\\website\\docs-website\\sidebars\\sidebars-solidworks-cpp.js",
        "sidebarCollapsible": true
      }
    ]
  ],
  "themeConfig": {
    "navbar": {
      "title": "The CAD Coder",
      "logo": {
        "alt": "My Site Logo",
        "src": "assets/logo.png"
      },
      "items": [
        {
          "label": "VBA",
          "position": "left",
          "to": "vba/vba-introduction"
        },
        {
          "label": "Solidworks API Tutorials",
          "position": "left",
          "type": "dropdown",
          "items": [
            {
              "label": "Solidworks API + VBA",
              "to": "/docs/intro"
            },
            {
              "label": "Solidworks API + C#",
              "to": "/docs/intro"
            },
            {
              "label": "Solidworks API + C++",
              "to": "solidworks-cpp/solidworks-Cpp-Api"
            }
          ]
        },
        {
          "label": "Resources",
          "position": "right",
          "to": "/resources"
        },
        {
          "label": "About Me",
          "position": "right",
          "to": "/aboutme"
        }
      ],
      "hideOnScroll": false
    },
    "footer": {
      "style": "light",
      "links": [
        {
          "title": "Docs",
          "items": [
            {
              "label": "Tutorial",
              "to": "/docs/intro"
            }
          ]
        },
        {
          "title": "Community",
          "items": [
            {
              "label": "Stack Overflow",
              "href": "https://stackoverflow.com/questions/tagged/docusaurus"
            },
            {
              "label": "Discord",
              "href": "https://discordapp.com/invite/docusaurus"
            },
            {
              "label": "Twitter",
              "href": "https://twitter.com/docusaurus"
            }
          ]
        },
        {
          "title": "More",
          "items": [
            {
              "label": "Blog",
              "to": "/blog"
            },
            {
              "label": "GitHub",
              "href": "https://github.com/facebook/docusaurus"
            }
          ]
        }
      ],
      "copyright": "Copyright © 2022 The CAD Coder."
    },
    "prism": {
      "additionalLanguages": [
        "csharp"
      ],
      "theme": {
        "plain": {
          "color": "#bfc7d5",
          "backgroundColor": "#292d3e"
        },
        "styles": [
          {
            "types": [
              "comment"
            ],
            "style": {
              "color": "rgb(105, 112, 152)",
              "fontStyle": "italic"
            }
          },
          {
            "types": [
              "string",
              "inserted"
            ],
            "style": {
              "color": "rgb(195, 232, 141)"
            }
          },
          {
            "types": [
              "number"
            ],
            "style": {
              "color": "rgb(247, 140, 108)"
            }
          },
          {
            "types": [
              "builtin",
              "char",
              "constant",
              "function"
            ],
            "style": {
              "color": "rgb(130, 170, 255)"
            }
          },
          {
            "types": [
              "punctuation",
              "selector"
            ],
            "style": {
              "color": "rgb(199, 146, 234)"
            }
          },
          {
            "types": [
              "variable"
            ],
            "style": {
              "color": "rgb(191, 199, 213)"
            }
          },
          {
            "types": [
              "class-name",
              "attr-name"
            ],
            "style": {
              "color": "rgb(255, 203, 107)"
            }
          },
          {
            "types": [
              "tag",
              "deleted"
            ],
            "style": {
              "color": "rgb(255, 85, 114)"
            }
          },
          {
            "types": [
              "operator"
            ],
            "style": {
              "color": "rgb(137, 221, 255)"
            }
          },
          {
            "types": [
              "boolean"
            ],
            "style": {
              "color": "rgb(255, 88, 116)"
            }
          },
          {
            "types": [
              "keyword"
            ],
            "style": {
              "fontStyle": "italic"
            }
          },
          {
            "types": [
              "doctype"
            ],
            "style": {
              "color": "rgb(199, 146, 234)",
              "fontStyle": "italic"
            }
          },
          {
            "types": [
              "namespace"
            ],
            "style": {
              "color": "rgb(178, 204, 214)"
            }
          },
          {
            "types": [
              "url"
            ],
            "style": {
              "color": "rgb(221, 221, 221)"
            }
          }
        ]
      },
      "magicComments": [
        {
          "className": "theme-code-block-highlighted-line",
          "line": "highlight-next-line",
          "block": {
            "start": "highlight-start",
            "end": "highlight-end"
          }
        }
      ]
    },
    "docs": {
      "sidebar": {
        "hideable": true,
        "autoCollapseCategories": false
      },
      "versionPersistence": "localStorage"
    },
    "colorMode": {
      "defaultMode": "light",
      "disableSwitch": false,
      "respectPrefersColorScheme": false
    },
    "metadata": [],
    "tableOfContents": {
      "minHeadingLevel": 2,
      "maxHeadingLevel": 3
    }
  },
  "themes": [
    [
      "D:\\website\\docs-website\\node_modules\\@easyops-cn\\docusaurus-search-local\\dist\\server\\server\\index.js",
      {
        "hashed": true
      }
    ]
  ],
  "baseUrlIssueBanner": true,
  "i18n": {
    "defaultLocale": "en",
    "locales": [
      "en"
    ],
    "localeConfigs": {}
  },
  "onDuplicateRoutes": "warn",
  "staticDirectories": [
    "static"
  ],
  "customFields": {},
  "scripts": [],
  "stylesheets": [],
  "clientModules": [],
  "titleDelimiter": "|",
  "noIndex": false
};