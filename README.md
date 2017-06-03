# cs-vsto-powerpoint-writer
Package provides VSTO-based C# implementation of a powerpoint modifier

# Objective

The project is to create a simple library which uses VSTO to automate powerpoint modification by using C# DSL syntax.

# Langauage

C# 

# Install

The library was built using VS2015 Community Edition. You can clone and build the library then add the library to your references in a .NET project. Note that this library is based on VSTO and thus requires the availability of office 2007 for it to work. It also requires the following COM libraries to be available in the C# project's references

* Microsoft.Office.Core (Version: 2.4)
* Microsoft.Office.Interop.Excel (Version: 1.6)
* Microsoft.Office.Interop.Powerpoint (Version: 2.9)

This link below shows how to solve the COM error when uninstall vs 2007 and reinstall some other version of office and then reinstall vs 2007:

https://social.msdn.microsoft.com/Forums/vstudio/en-US/08f13e9d-895c-4102-b6d9-e327af8cf8c0/0x80029c4a-typeecantloadlibrary?forum=vsto

# Usage

Below is the C# sample code for creating a sample report:

```cs 

```

