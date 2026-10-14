#!/usr/bin/env bash
# Genera manifest.dev.xml apuntando a https://localhost:3000 con un Id distinto.
set -euo pipefail
cd "$(dirname "$0")/.."
sed \
  -e 's#https://xictorlr.github.io/skillsEEpowerPoint/#https://localhost:3000/#g' \
  -e 's#<AppDomain>https://xictorlr.github.io</AppDomain>#<AppDomain>https://localhost:3000</AppDomain>#' \
  -e 's#<Id>7b2c1f3e-5a84-4d6b-9c1e-2f8a4b6d9e01</Id>#<Id>7b2c1f3e-5a84-4d6b-9c1e-2f8a4b6d9e02</Id>#' \
  -e 's#DisplayName DefaultValue="Slide Tools"#DisplayName DefaultValue="Slide Tools (dev)"#' \
  manifest.xml > manifest.dev.xml
echo "manifest.dev.xml generado"
