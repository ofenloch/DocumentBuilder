# DocumentBuilder

Build DOCX and XLSX documents from templates with C\#, .NET and [Open XML SDK](https://docs.microsoft.com/en-us/office/open-xml/open-xml-sdk).


The project's structure looks like this

        .
        ├── data
        ├── DocumentBuilder
        │   ├── DocumentBuilder.csproj
        │   ├──  source files
        ├── DocumentBuilder.sln
        ├── lib
        │   ├── dblib.csproj
        │   ├──  source files 
        ├── output
        ├── README.md (this file)
        └── scripts
        └── test
            ├── unit-tests.csproj
            └──   source files

**Directory DocumentBuilder** contains the sources for application/program/binary. I would 
have called this directory bin/ but that is already used.

**Directory lib** contains the sources for dblib, the library used in the application.

**Directory data** contains data needed for hte project, e.g. logger configuration, 
(test) input data and things like that.

The unit test project and its source files are locate in **directory test**.

Deploymen and other scripts reside in **directory scripts**.

The tests and launch configurations write the outputfiles innto **directory output**. 
This directory is listed in file *.gitignore* and can be safely removed. All the  *bin/* 
and *obj* directories are ignored by Git and can be safely removed. See script *scripts/make-clean.sh*.


## Re-Organize Project

create directory DocumentBuilder and move Program.cs and DocumentBuilder.csproj there

create a new solution file with `dotnet new sln`

add project DocumentBuilder to solution: `dotnet sln add DocumentBuilder/`

add project dblib to solution: `dotnet sln add lib/`

edit file .vscode/tasks.json and use the solution file *DocumentBuilder.sln* for the build and publish tasks

See [my project "hello-world"](https://github.com/ofenloch/hello-world) for a detailed explanation 
about re-structuring the project, adding unit tests and logging.

## References

https://docs.microsoft.com/en-us/office/open-xml/packages-and-general

https://docs.microsoft.com/en-us/office/open-xml/spreadsheets

https://docs.microsoft.com/en-us/office/open-xml/word-processing
