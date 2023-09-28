# Office File API - Use Office File API Libraries as a Back-End in Web-API Applications

This example shows how to create an ASP.NET Core Web API application with the Office File API libraries. The application contains a series of endpoints to perform the following actions:

- Convert Word and Excel files to available formats
- Split Word, Excel and PDF files
- Merge Word, Excel and PDF files
- Password-protect Word, Excel and PDF files

 The project contains a Dockerfile that specifies how to build the application in a Docker container.

 > [!note]
 > You need a license for the [DevExpress Office File API](https://www.devexpress.com/products/net/office-file-api/) Subscription or [DevExpress Universal](https://www.devexpress.com/subscriptions/universal.xml) Subscription to use Office File API libraries in production code.

## Files to Review

 * [PdfController.cs](./CS/Controllers/PdfController.cs)
 * [RichEditController.cs](./CS/Controllers/RichEditController.cs)
 * [SpreadsheetController.cs](./CS/Controllers/SpreadsheetController.cs)

## More Examples

* [How to Dockerize an Office File API Application](https://github.com/DevExpress-Examples/dockerize-office-file-api-app)