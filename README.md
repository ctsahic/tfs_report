# Azure DevOps PR Report

A Python utility to generate pull request reports from Azure DevOps, with support for date ranges, repository filtering, and work item associations.

## Features

- 📊 Generate PR reports with comprehensive metadata (status, merge status, work items, etc.)
- 🗓️ Filter pull requests by date and time ranges
- 📁 Support for single or multiple repository selection
- 💾 Multiple export formats: Excel, HTML, and JSON
- 🖥️ Graphical User Interface (GUI) for easy access
- 💻 Command-line interface for automation and scripting
- 🔐 Secure PAT (Personal Access Token) management with local caching
- 📝 Auto-loading of saved repositories for quick access

## Requirements

- Python 3.9+
- Azure DevOps Personal Access Token (PAT)

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/ctsahic/tfs_report.git
   cd tfs_report
