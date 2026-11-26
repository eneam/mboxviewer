# Build baseline

## Pré-requisitos

- Visual Studio 2022 Build Tools (ou Visual Studio 2022 Community) com:
  - `Desktop development with C++` (`Microsoft.VisualStudio.Workload.VCTools`)
  - `MSVC v143 - VS 2022 C++ x64/x86 build tools`
  - `C++ MFC for latest v143 build tools (x86 & x64)`
  - `Windows 11 SDK (10.0.26100)` ou outro SDK `10.0.x`

## Comando validado

```powershell
& 'C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\amd64\MSBuild.exe' 'mboxview.sln' /p:Configuration=Release /p:Platform=x64 /v:minimal /nologo /m
```

Saída esperada:

```text
mboxview.vcxproj -> x64\Release\mboxview.exe
```

## Estado verificado

- `Release|x64` compila a partir do código-fonte.
- O executável inicia e permanece ativo após o arranque (smoke test).
- Existe um warning pré-existente `C4244` em `src/FileConfigurationDB.cpp`; não bloqueia o build.

## Nota de instalação

O `winget` pode reportar `0x8013153B` mesmo quando o instalador do Visual Studio concluiu a instalação.
Se isso acontecer, confirmar diretamente a presença de `MSBuild.exe`, `cl.exe`, `afxwin.h` e do SDK antes de repetir a instalação.
