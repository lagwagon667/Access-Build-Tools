# Access Build Tools

![logo](resources/logo.png)

# 📦 Access Build Tools

A PowerShell module for **building Microsoft Access databases from source**.  
It is designed to support a full build pipeline including versioning, publishing, testing, and ACCDE compilation.

---

## ⚙️ Requirements

To use the build script, the following components must be installed on the system:

- **Microsoft Access**
- **[JoyfullService/ms-access-vcs Add-In](https://github.com/joyfullservice/msaccess-vcs-addin)** [v5.1-alpha or newer](https://github.com/lagwagon667/msaccess-vcs-addin/releases/download/v5.1.0-alpha-1/Version.Control.zip)
  (used for exporting/importing Access objects as source files)

---

## 📥 Installation

You can install the module in the following way:

### 🔹 Local Installation (recommended for development)

Place the `AccessBuildTools` folder into one of your PowerShell module paths, e.g.:

```
$env:USERPROFILE\Documents\PowerShell\Modules\
```

Then import it:

```powershell
Import-Module AccessBuildTools
```

---

## 🚀 CmdLets

- **Build-Accdb** — Creates an Access database from source files.  
  Optionally generates a version number and injects it into the database during the build process.

- **Publish-Accdb** — Embeds external references directly into the database and injects startup code that unpacks these references before Access launches the application.  
  This ensures the database runs correctly on machines where the referenced libraries are not installed.

- **Invoke-UnitTests** — Executes automated tests defined in the Access project to validate functionality during CI/CD or local development.

- **Convert-ToAccde** — Produces a compiled ACCDE file for deployment, ensuring code protection and stable runtime behavior.

- **Get-References** — Outputs all **not-builtin** references used by a given Access database.  
  This utility helps determine which references should be embedded when running the `Publish-Accdb` cmdlet.

For further information see the builtin help of the CmdLets (e.g. `help Build-Accdb`)

## 📦 Typical Workflows

Because **Publish-Accdb modifies the Access DB**, it cannot be combined with ACCDE compilation.  
A build pipeline therefore splits into two valid paths:

### 🔧 Workflow 1 — Build → Test → Publish

Use this workflow when you want an **ACCDB with embedded references** that can run on machines without the external libraries installed.

1. **Build-Accdb**
2. **Invoke-UnitTests**
3. **Publish-Accdb**  
   (Embeds references and injects unpack/startup code)

### 🛡️ Workflow 2 — Build → Test → Compile

Use this workflow when you want a **final ACCDE** for deployment.

1. **Build-Accdb**
2. **Invoke-UnitTests**
3. **Convert-ToAccde**  
   (Produces the compiled ACCDE)
