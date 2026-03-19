# 🧩 Terraform AzureRM Versioned Schema Extraction Guide

This guide walks through creating a **version-specific Terraform workspace** and exporting a **complete AzureRM provider schema** along with a **clean list of resource types**.

---

## 📁 1. Create a Version-Specific Directory

```bash
mkdir Terraform_v4_XX_0
cd Terraform_v4_XX_0
```

---

## 📄 2. Create the Terraform Configuration File

```bash
touch main.tf
```

---

## ✍️ 3. Add Version-Specific Configuration

Append the following content to `main.tf`:

```hcl
terraform {
  required_providers {
    azurerm = {
      source  = "hashicorp/azurerm"
      version = "4.XX.0"
    }
  }
}

provider "azurerm" {
  features {}
}
```

> 🔁 Replace `4.XX.0` with your desired AzureRM provider version (e.g., `4.53.0`)

---

## ⚙️ 4. Initialize Terraform

```bash
terraform init -input=false
```

This step:
- Downloads the specified AzureRM provider version
- Prepares the working directory

---

## 📤 5. Export Full Provider Schema (JSON)

```bash
terraform providers schema -json > azurerm-4.XX.0-schema.json
```

This file contains:
- All resource definitions
- All data sources
- Full schema metadata for the specified version

---

## 🔍 6. Extract Resource Types (Using `jq`)

### ✅ Prerequisites
- Install `jq` (JSON processor)

#### macOS (Homebrew)
```bash
brew install jq
```

---

### 🧠 Generate Clean Resource Type List

```bash
jq -r '
  .provider_schemas["registry.terraform.io/hashicorp/azurerm"]
  .resource_schemas
  | keys
  | sort
' azurerm-4.XX.0-schema.json > azurerm-4.XX.0-resource-types.json
```

---

## 📦 Output Files

| File Name | Description |
|----------|-------------|
| `azurerm-4.XX.0-schema.json` | Full provider schema |
| `azurerm-4.XX.0-resource-types.json` | Sorted list of all resource types |

---

## 🚀 Example Output (Resource Types)

```json
[
  "azurerm_api_management",
  "azurerm_application_gateway",
  "azurerm_resource_group",
  "azurerm_storage_account"
]
```

---

## 💡 Tips

- Use a new directory per version to avoid conflicts  
- Keep schema files for **version diffing** and governance  
- Useful for:
  - Sentinel policies
  - Compliance checks
  - Resource validation pipelines  

---

## 🧠 Optional Enhancements

- Extract **data sources** similarly using `.data_source_schemas`
- Convert JSON → CSV for reporting
- Diff multiple versions to track provider changes

---

Happy building! 🚀
