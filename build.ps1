function Get-ImageTag {
    $latestTag = git rev-list --tags --max-count=1
    if ($latestTag) {
        $tagName = git describe --tags $latestTag
        $shortHash = git rev-parse --short HEAD
        return "$tagName-$shortHash"
    } else {
        Write-Error "No git tags found. Please create a tag first with: git tag v1.0.0"
        exit 1
    }
}

function Build-Image {
    param (
        [string]$environment = "default"
    )
    
    $imageTag = Get-ImageTag
    
    if ($environment -eq "ppi") {
        $registry = "asia-southeast2-docker.pkg.dev/eternal-skyline-166605/bpcs-image-registry"
    } else {
        $registry = "asia-southeast2-docker.pkg.dev/data-commstrexe-prd-565x/bpcs-image-registry"
    }
    
    $fullImageName = "$registry/itcomm-addin:$imageTag"
    Write-Host "Building image: $fullImageName"
    docker build . --platform linux/amd64 -t $fullImageName
}

function Push-Image {
    param (
        [string]$environment = "default"
    )
    
    $imageTag = Get-ImageTag
    
    if ($environment -eq "ppi") {
        $registry = "asia-southeast2-docker.pkg.dev/eternal-skyline-166605/bpcs-image-registry"
    } else {
        $registry = "asia-southeast2-docker.pkg.dev/data-commstrexe-prd-565x/bpcs-image-registry"
    }
    
    $fullImageName = "$registry/itcomm-addin:$imageTag"
    Write-Host "Pushing image: $fullImageName"
    docker push $fullImageName
}

# Get the command line arguments
$command = $args[0]

switch ($command) {
    "build" { Build-Image }
    "build-ppi" { Build-Image -environment "ppi" }
    "push" { Push-Image }
    "push-ppi" { Push-Image -environment "ppi" }
    default {
        Write-Host "Usage: .\build.ps1 <command>"
        Write-Host "Commands:"
        Write-Host "  build      - Build the default image"
        Write-Host "  build-ppi  - Build the PPI image"
        Write-Host "  push       - Push the default image"
        Write-Host "  push-ppi   - Push the PPI image"
    }
}