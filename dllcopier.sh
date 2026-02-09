#!/bin/bash

WINCROSSBUILD=${1}
EXEPATH=${2}
echo "Using Windows cross build path: ${WINCROSSBUILD}"

DLLs=(
    "libstdc++-6.dll"
    "libgcc_s_seh-1.dll"
    "libwinpthread-1.dll"
    "zlib1.dll"
    "libMiniXLSX.dll"
    "libKFZippa.dll"
)

DLL_SearchPath=(
    "/usr/x86_64-w64-mingw32/bin"
    "${WINCROSSBUILD}"
    "${WINCROSSBUILD}/KFZippa"
    "${WINCROSSBUILD}/MiniXLSX"
)

EXE_FOLDER=$(dirname "${EXEPATH}")

if [ ! -d "${EXE_FOLDER}" ]; then
    echo "Creating executable folder: ${EXE_FOLDER}"
    mkdir -p "${EXE_FOLDER}"
fi

for dll in "${DLLs[@]}"; do
    found=0
    for path in "${DLL_SearchPath[@]}"; dogit
        # echo "Searching for ${dll} in ${path}"
        if [ -f "${path}/${dll}" ]; then
            # echo "Copying ${dll} from ${path}"
            cp "${path}/${dll}" "${EXE_FOLDER}" -rv
            found=1
            break
        fi
    done
    if [ $found -eq 0 ]; then
        echo "Error: ${dll} not found in any search path."
        exit 1
    fi
done