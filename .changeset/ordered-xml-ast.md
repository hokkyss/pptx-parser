---
"@hokkyss/pptx-core": minor
"@hokkyss/pptx-reader": minor
"@hokkyss/pptx-writer": minor
"@hokkyss/pptx": minor
---

refactor: migrate write and read engines to 1st-party array-based `XmlElement` AST

- **Zero External XML Dependencies**: Completely removed `fast-xml-parser` from the monorepo, replacing it with a custom zero-dependency isomorphic XML serializer and streaming recursive-descent parser.
- **Ordered `XmlElement` AST in `@hokkyss/pptx-core`**: Introduced canonical, strongly-typed `XmlElement`, `XmlChild`, `XmlRaw`, and `XmlParser` interfaces shared across reader and writer.
- **Sequential Document & Z-Order Preservation**:
  - Maintained exact visual z-order for shape trees (`<p:spTree>`, `<p:grpSp>`) across shapes, pictures, connectors, and graphic frames without tag-clustering.
  - Maintained exact interleaving order for paragraph children (`<a:p>`), supporting text runs (`<a:r>`), fields (`<a:fld>`), and soft line breaks (`<a:br>`).
- **Full Type-Safety in Serializers**: All writer serializer modules now return strictly-typed `XmlElement` AST nodes instead of loosely-typed object trees.
