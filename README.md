<!-- default badges list -->
![](https://img.shields.io/endpoint?url=https://codecentral.devexpress.com/api/v1/VersionRange/697761947/23.1.5%2B)
[![](https://img.shields.io/badge/Open_in_DevExpress_Support_Center-FF7200?style=flat-square&logo=DevExpress&logoColor=white)](https://supportcenter.devexpress.com/ticket/details/T1192524)
[![](https://img.shields.io/badge/📖_How_to_use_DevExpress_Examples-e9f6fc?style=flat-square)](https://docs.devexpress.com/GeneralInformation/403183)
<!-- default badges end -->
# Office File API - Use DevExpress Office File API Libraries (at the backend) for your Web-API Applications

By following the steps outlined in this example, you’ll create an ASP.NET Core Web API application using DevExpress Office File API libraries. The application will include a series of endpoints designed to perform the following actions:

* Convert Word and Excel files to available formats.
* Split Word, Excel and PDF files.
* Merge Word, Excel and PDF files.
* Password-protect Word, Excel and PDF files.

The project contains a Dockerfile that specifies how to build the application in a Docker container.

## Build the Docker Image

Obtain a [DevExpress NuGet Feed URL](https://docs.devexpress.com/GeneralInformation/116042/installation/install-devexpress-controls-using-nuget-packages/obtain-your-nuget-feed-credentials) and add your feed authorization key to the [Dockerfile](./CS/Dockerfile) to specify the package source for the [DevExpress.Document.Processor](https://nuget.devexpress.com/packages/DevExpress.Document.Processor/) NuGet package.

Use the following commands to build and run the docker image:

  ```
  docker build -t officefileapinwebapi .
  docker run -d -p 8080:80 officefileapinwebapi
  ```


> [!Note] 
> You need to purchase a license to use the DevExpress Office File API in production code (DevExpress Office File API Subscription or DevExpress Universal Subscription)

## Files to Review

 * [PdfController.cs](./CS/Controllers/PdfController.cs)
 * [RichEditController.cs](./CS/Controllers/RichEditController.cs)
 * [SpreadsheetController.cs](./CS/Controllers/SpreadsheetController.cs)

## More Examples

* [How to Dockerize an Office File API Application](https://github.com/DevExpress-Examples/dockerize-office-file-api-app)
