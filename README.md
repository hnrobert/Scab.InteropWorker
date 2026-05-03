# Scab.InteropWorker

SVN CAD Asset Browser interop worker. Windows-only .NET 10 process that performs CAD file rendering.

## Overview

InteropWorker runs as a separate process, exposing a MagicOnion gRPC server on `127.0.0.1:5100`. ServerD dispatches render jobs to it.

## Supported Job Types

| Type | Description |
|------|-------------|
| `InventorPng` | Export Autodesk Inventor file (.ipt/.iam/.idw) to PNG via COM automation |
| `StepRender` | Headless render STEP file to PNG |
| `StlRender` | Headless render STL file to PNG |

## Requirements

- Windows OS
- Autodesk Inventor installed (for InventorPng jobs)
- The worker must be able to create a COM instance of `Inventor.Application`

## Job Protocol

1. ServerD calls `SubmitJobAsync` with file path + parameters
2. Worker opens file in Inventor (or renderer)
3. Exports rendered view as PNG
4. Returns output path to ServerD
5. ServerD stores PNG in thumbnail cache

## Running

```bash
dotnet run --project src/Scab.InteropWorker [port]
# Default port: 5100
```

ServerD can also auto-start the worker process by configuring `Scab:WorkerExePath` in `appsettings.json`.

## COM Interop

Inventor COM automation uses late binding via `System.Runtime.InteropServices.Marshal.GetActiveObject("Inventor.Application")`. No primary interop assemblies required.
