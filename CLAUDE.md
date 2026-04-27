# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Architecture

This project uses Contensive patterns and follows the contensive architecture and patterns
- [Contensive Architecture](https://raw.githubusercontent.com/contensive/Contensive5/refs/heads/master/README.md)

## Project Overview

This is a **web spider/crawler addon** built on the **Contensive CMS framework** (.NET Standard 2.0). It crawls website links, extracts content (text, images, OpenGraph metadata), and stores the data for full-text search capabilities.

### Contensive Framework Context

This project operates as an **Extension Layer** addon within the Contensive5 framework's 5-tier architecture:

1. **Presentation Layer** - AdminSite addon and WebApi (ASP.NET Core)
2. **Extensions Layer** - **This Spider addon lives here**
3. **Processing Core** - CoreController and Processor orchestrating operations
4. **Data Management** - Database entity definitions
5. **Foundation** - CPBase abstract classes with infrastructure services

Reference: [Contensive5 Repository](https://github.com/contensive/Contensive5)

### Contensive Design Principles

The framework emphasizes:

1. **Hardware Abstraction** - `CPBaseClass` (cp object) abstracts all infrastructure:
   - Database access: `cp.Db`, `cp.Content`
   - Caching: `cp.Cache` (Redis)
   - File system: `cp.CdnFiles`, `cp.PrivateFiles`
   - HTTP context: `cp.Doc`, `cp.Response`
   - Email: `cp.Email`

2. **Modular Reusability** - Everything is an addon:
   - Installed via collection XML packages
   - Easy to compose, test, and reuse across projects

3. **Metadata-Driven** - Data structure defined in metadata, not migrations:
   - Content Definitions describe entities
   - Auto-generates admin UI for CRUD operations

4. **Testability** - Dependency injection via cp object

### Contensive Architecture and Patterns

- [Contensive Architecture](https://github.com/contensive/Contensive5/blob/master/patterns/contensive-architecture.md)

## Build Commands

**Main Build:**
```cmd
cd scripts
build.cmd
```
This builds the project, packages the collection ZIP, and deploys to `C:\Deployments\aoSpider\Dev\{version}`.

**Build Individual Project:**
```cmd
cd server
dotnet build Spider/Spider.csproj --configuration Debug
```

## Solution Structure

- **server/Spider/** - Main spider implementation (crawling, content extraction, cleanup)
- **collections/aoSpider/** - Collection definition XML and deployment package
- **scripts/** - Build automation

## Contensive Platform Patterns

### Add-on Execution Model

All features are "Add-ons" - classes that inherit `AddonBaseClass` and implement `Execute(CPBaseClass cp)`:

```csharp
public class SpiderClass : AddonBaseClass {
    public override object Execute(CPBaseClass cp) {
        // cp provides typed framework access
        // Feature implementation
        return result;
    }
}
```

### Database/ORM Pattern

```csharp
public class LinkAliasModel : DbBaseModel {
    public static DbBaseTableMetadata tableMetadata { get; } = new DbBaseTableMetadata {
        contentName = "Link Aliases",
        tableName = "ccLinkAliases"
    };
    public int pageid { get; set; }
    public string querystringsuffix { get; set; }
}
```

### Collection Installation Pattern

`collections/aoSpider/Spider.xml` defines the installation package:
- Addon definitions with GUIDs and DotNet class mappings
- Content Definitions for database tables (Spider Docs, Spider Links, etc.)
- Collection ZIP contains: XML + DLLs

## Key Addons

1. **Spider** (`SpiderClass`) - Main crawler that fetches pages, extracts content, updates Spider Docs
2. **Spider All Unspidered Links** (`SpiderAllUnspideredLinks`) - Batch processes all unspidered links
3. **Spider Clean Up** (`SpiderCleanUp`) - Removes orphaned Spider Docs entries

## Code Standards

- Use string interpolation over concatenation
- In C# add curly braces around nested statements
- Follow existing Contensive naming conventions
- Use `cp.Site.ErrorReport(ex)` for exception logging
