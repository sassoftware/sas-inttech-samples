# Building Windows applications with SAS Integration Technologies

These are example applications for Windows developers to integrate SAS with Microsoft.NET and PowerShell, using SAS Integration Technologies. You may use them as-is or incorporate into your own applications to integrate with SAS 9.x software. Unless otherwise noted, all examples should work with SAS 9.2 and later.

This repository is a repackaging of examples that support this paper: [Create Your Own Client Apps using SAS Integration Technologies](http://support.sas.com/resources/papers/proceedings13/003-2013.pdf).

## Overview

SAS provides an open programming interface for Windows developers to connect to SAS 9 environments. This open interface allows you to run SAS programs, access data, upload and download files, and more. The APIs are surfaced through the SAS Integration Technologies product. If you use SAS Enterprise Guide or SAS Add-In for Microsoft Office in your organization, then you probably already have the infrastructure that you need to try these examples.

These examples include:

* [SASHarness](Microsoft.NET/SASHarness) -- a simple Microsoft .NET application (C#) that connects to a SAS Workspace, submits code, retrieves the SAS log and listing. It also includes a simple SAS data viewer.
* [Windows PowerShell examples (various)](PowerShell/) -- PowerShell scripts that connect to SAS Metadata Server, SAS Workspace, and local SAS data sets.

### Prerequisites

These examples require Microsoft Windows (any modern version, workstation or server). They also rely on an installed component from SAS called "SAS Integration Technologies client". This component is [available as a free download from support.sas.com](https://support.sas.com/downloads/browse.htm?fil=&cat=56). If you have no other SAS applications on your Windows machine, you will probably need to download/install this component. If you use SAS applications such as SAS for Windows, SAS Enterprise Guide, or SAS Add-In for Microsoft Office, then you should already have the SAS Integration Technologies client.

To explore and build Microsoft.NET examples, you'll need Microsoft Visual Studio. The free community edition is adequate for opening and building these examples.

For the PowerShell examples, you'll need access to Windows PowerShell. It's built into the Microsoft Windows environment -- but the [ability to create and run your own scripts might be restricted in your organization](https://blogs.sas.com/content/sasdummy/2011/09/12/running-windows-powershell-scripts/).

## Contributing

We welcome your contributions! Please read [CONTRIBUTING.md](CONTRIBUTING.md) for details on how to submit contributions to this project.

## License

This project is licensed under the [Apache 2.0 License](LICENSE).

## Additional Resources

Helpful information about how to use SAS Integration Technologies and its APIs is available in SAS Communities, blogs, and documentation.

* [SAS Communities article that describes the process](https://communities.sas.com/t5/SAS-Communities-Library/Create-your-own-client-apps-using-SAS-Integration-Technologies/ta-p/418253)
* Désilets, Karine. 2012. ["SAS IOM and Your .NET Application Made Easy"](http://support.sas.com/resources/papers/proceedings12/017-2012.pdf), Proceedings of the SAS Global Forum 2012 Conference
* [SAS and Windows PowerShell](https://blogs.sas.com/content/sasdummy/tag/powershell/) blog series
* [SAS Integration Technologies: Windows Client Developer's Guide](https://go.documentation.sas.com/?docsetId=itechwcdg&docsetTarget=p0xyn6hf6w4e0an1dum74ovbcphw.htm&docsetVersion=9.4&locale=en)
