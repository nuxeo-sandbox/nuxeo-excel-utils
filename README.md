# nuxeo-excel-utils

> [!IMPORTANT]
> This is **W**ork **I**n **P**rogress, using GotHub as backup, for a POC


## Features

_Tp Be Done and Continued_

### Automation
#### `Excel.GetProperties`(in `Files`)
* input
  * `Document` or `Blob`
  * If `Document`, the optional `xpath` parameter is used to get the blob (default value is `file:content`)
* output
  * A JSONBlob, holding an object with the following sub-objects:
  * `sheets`: An array with the names of the sheets
  * `coreProperties`: An object with the core properties (title, creator, category, ...). See source code for all possible values
  * `customProperties`: An object with the custom properties. **WARNING**: ONly String properties are returned with a value, all other are returned as `null` even if not null.


## Build and Install

  ```
  cd /path/to/nuxeo-excel-utils
  mvn clean install
  ```


## Licensing

[Apache License, Version 2.0](http://www.apache.org/licenses/LICENSE-2.0)


## Support
These features are not part of the Nuxeo Production platform.

These solutions are provided for inspiration and we encourage customers to use them as code samples and learning resources.

This is a moving project (no API maintenance, no deprecation process, etc.) If any of these solutions are found to be useful for the Nuxeo Platform in general, they will be integrated directly into platform, not maintained here.


## About Nuxeo
Nuxeo Platform is an open source Content Services platform, written in Java and provided by Hyland. Data can be stored in both SQL & NoSQL databases.

The development of the Nuxeo Platform is mostly done by Nuxeo employees with an open development model.

The source code, documentation, roadmap, issue tracker, testing, benchmarks are all public.

Typically, Nuxeo users build different types of information management solutions for [document management](https://www.nuxeo.com/solutions/document-management/), [case management](https://www.nuxeo.com/solutions/case-management/), and [digital asset management](https://www.nuxeo.com/solutions/dam-digital-asset-management/), use cases. It uses schema-flexible metadata & content models that allows content to be repurposed to fulfill future use cases.

More information is available at [www.hyland.copm](https://www.hyland.com).
