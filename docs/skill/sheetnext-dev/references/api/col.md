# Col API

> Generated: 2026-09-08
> SheetNext source package: 0.2.23; source fingerprint: 896bd17c14ed91f3295183303ef16d4d78bb335764d4362f17ab30915bafee95

[API index](../core-api.md)

- [Col](#col)

## Col
- Source: `src/core/Col/Col.js`
- Column object

### Props
| Name | Type | Static | Default | Description |
| --- | --- | --- | --- | --- |
| sheet | Sheet | No | sheet | Sheet it belongs to |
| cIndex | number | No | cIndex | Column Index |

### Get/Set
| Name | Type | Mode | Static | Description |
| --- | --- | --- | --- | --- |
| width | number | get/set | No | Column Width (px) |
| hidden | boolean | get/set | No | Whether to hide |
| outlineLevel | number | get/set | No | Outline Level (0-7) |
| collapsed | boolean | get/set | No | Outline Collapse Marker |
| style | Object | get/set | No | Column Styles |
| font | Object | get/set | No | Font style |
| fill | Object | get/set | No | Fill Style |
| alignment | Object | get/set | No | Alignment Style |
| border | Object | get/set | No | Border Style |
| numFmt | string \| undefined | get/set | No | Number Format |

### Get
| Name | Type | Static | Description |
| --- | --- | --- | --- |
| cells | Array<Cell> | No | Sparse cells indexed by row |
