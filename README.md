# Export Data to Excel in Blazor Application

This repository demonstrates how to use Syncfusion Excel library (Essential XlsIO) in Blazor applications to create and export Excel files without Microsoft Office dependencies. It provides examples for both client-side Blazor applications and server-side Blazor applications, showing how to generate Excel documents dynamically and download them through the browser.

In the client-side application, the sample illustrates creating an Excel workbook directly within the Blazor component. A DataTable is initialized with sample data such as salespersons, ages, and salaries, and then imported into the worksheet using worksheet.ImportDataTable. The worksheet applies auto-fit operations to columns for clean formatting. The workbook is saved into a memory stream and downloaded to the client using JavaScript interop (JS.SaveAs). This approach demonstrates how Blazor WebAssembly applications can generate Excel files entirely on the client side, making them lightweight and independent of server processing.

In the server-side application, the sample uses a service (ExcelService) to create the Excel document on the server. The Blazor component calls this service, retrieves the generated workbook as a memory stream, and then downloads the file to the client using JavaScript interop. This approach is useful when Excel generation requires server-side resources or integration with backend systems. The server-side example highlights how Blazor Server applications can seamlessly integrate Syncfusion XlsIO for Excel creation and provide the file to users through the browser.

Both approaches showcase the versatility of Syncfusion XlsIO in Blazor, enabling developers to create, read, edit, and convert Excel files without relying on Microsoft Office. The repository demonstrates how to integrate Excel export functionality into modern Blazor applications, whether running on the client or server. By following these examples, developers can adapt the workflow to their own scenarios, ensuring that Excel reports are generated dynamically with consistent formatting and delivered directly to end users.

## Prerequisites

* Visual Studio 2022

## How to run the project

* Checkout this project to a location in your disk.
* Open the solution file using the Visual Studio 2022.
* Restore the NuGet packages by rebuilding the solution.
* Run the project.