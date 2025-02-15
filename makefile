IMAGE_NAME := $$(git describe --tags $$(git rev-list --tags --max-count=1))-$$(git rev-parse --short HEAD)
build:
        docker build . --platform linux/amd64 -t asia-southeast2-docker.pkg.dev/data-commstrexe-prd-565x/bpcs-image-registry/itcomm-addin:$(IMAGE_NAME)

build-ppi:
        docker build . --platform linux/amd64 -t asia-southeast2-docker.pkg.dev/eternal-skyline-166605/bpcs-image-registry/itcomm-addin:$(IMAGE_NAME)

push:
        docker push asia-southeast2-docker.pkg.dev/data-commstrexe-prd-565x/bpcs-image-registry/itcomm-addin:$(IMAGE_NAME)

push-ppi:
        docker push asia-southeast2-docker.pkg.dev/eternal-skyline-166605/bpcs-image-registry/itcomm-addin:$(IMAGE_NAME)