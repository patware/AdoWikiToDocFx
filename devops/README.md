# Azure DevOps Pipeline

Given a fictional scenario where you have the following DocFx folder/file structure

- Root/
  - index.md
    - Products/
      - index.md
      - SomeApp/
      - Acme/
    - APIs
      - index.md
      - Foo/
      - Bar/

## AdoWikiDocFxHelper.json

The corresponding AdoWikiDocFxHelper.json file would be something like:

```json
{
  "wikis" : [
    {
      "name": "Wiki_Main",
      "cloneUrl": "clone url of the wiki ex: https://fabrikam@dev.azure.com/fabrikam/Documentation/_git/Wiki.Main",
      "wikiUrl": "https://dev.azure.com/fabrikam/Documentation/_wiki/wikis/Wiki.Main"
    },
    {
      "name": "Wiki_SomeApp",
      "target": "/Products/SomeApp",
      "menuDisplayName": "Some App",
      "menuPosition": "1",
      "cloneUrl": "https://fabrikam@dev.azure.com/fabrikam/SomeProject/_git/Wiki.SomeApp",
      "wikiUrl": "https://dev.azure.com/fabrikam/SomeProject/_wiki/wikis/Wiki.SomeApp"
    },
    {
      "name": "Wiki_Acme",
      "target": "/Products/Acme",
      "menuDisplayName": "Acme",
      "menuPosition": "0",
      "cloneUrl": "https://fabrikam@dev.azure.com/fabrikam/Acme/_git/Wiki.Acme",
      "wikiUrl": "https://dev.azure.com/fabrikam/Acme/_wiki/wikis/Wiki.Acme",
      "wikiRootFolder": "documentation"
    }
  ],
  "apis" : [
    {
      "cloneUrl": "https://dev.azure.com/fabrikam/Common/_git/Foo",
      "code" : "src/"
    },
    {
      "cloneUrl": "https://dev.azure.com/fabrikam/Common/_git/Bar",
      "code" : "src/"
    }
  ]
}
```

### Wiki properties

| Property        | Type              | Description                                                       | Default            | Example                                                               |
|-----------------|-------------------|-------------------------------------------------------------------|--------------------|-----------------------------------------------------------------------|
| name            | string            | Name of the wiki repo in the pipeline                             |                    | Wiki_SomeApp                                                          |
| cloneUrl        | string            | Git CloneUrl of the wiki                                          |                    | https://fabrikam@dev.azure.com/fabrikam/SomeProject/_git/Wiki.SomeApp |
| wikiUrl         | string            | Wiki's root URL                                                   |                    | https://dev.azure.com/fabrikam/SomeProject/_wiki/wikis/Wiki.SomeApp   |
| target          | [optional]string  | sub folder path where the wiki will end up                        | /                  | /Products/SomeApp                                                     |
| menuDisplayName | [optional]string  | Title of the page that will be displayed in the TOC               | wiki's folder name | Some App                                                              |
| menuPosition    | [optional]integer | (zero-based) Position where the entry will be inserted in the TOC.  Added to the end of the items if null | null               | 2                                                                     |
| wikiRootFolder  | [optional]string  | sub folder where the actual wiki pages are located                | null               | documentation                                                         |

Notes on the example above.

The SomeApp Wiki is a child wiki, and will end up under /Products/SomeApp.  (property: target)

The /Products/toc.yml will have an entry with display: "Some App", and will be the 3rd entry in the products's TOC (properties menuDisplayName and menuPosition):

1. Some entry
1. Another entry
1. Some App -> 3rd entry
1. Bla bla

The Ado Wiki is a "Publish code as Wiki", and the wiki pages aren't at the root folder, but in a subfolder named "documentation" (property wikiRootFolder)
