# DocumentBuilder

Build DOCX and XLSX documents from templates with C\#, .NET and [Open XML SDK](https://docs.microsoft.com/en-us/office/open-xml/open-xml-sdk).





## Re-Organize Project

create directory DocumentBuilder and move Program.cs and DocumentBuilder.csproj there

create a new solution file with `dotnet new sln`

add project DocumentBuilder to solution: `dotnet sln add DocumentBuilder/`

add project dblib to solution: `dotnet sln add lib/`

edit file .vscode/tasks.json and use the solution file *DocumentBuilder.sln* for the build and publish tasks

(see [my project "hello-world"](https://github.com/ofenloch/hello-world) for a detailed explanation)

## References

https://docs.microsoft.com/en-us/office/open-xml/packages-and-general

https://docs.microsoft.com/en-us/office/open-xml/spreadsheets

https://docs.microsoft.com/en-us/office/open-xml/word-processing
