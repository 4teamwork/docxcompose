/*
# To build locally use:
GIT_TAG=$(git describe --tags --candidates=0) \
GIT_SHA_TAG=$(git describe --tags) \
LATEST_TAG=$(git describe --tags --abbrev=0 master) \
BRANCH_NAME=$(git rev-parse --abbrev-ref HEAD) \
docker buildx bake -f docker-bake.hcl --load
*/

variable "IMAGE_NAME" {
  default = "docker.io/4teamwork/docxcompose"
}
variable "GIT_TAG" {
  default = ""
}
variable "GIT_SHA_TAG" {
  default = ""
}
variable "LATEST_TAG" {
  default = ""
}
variable "BRANCH_NAME" {
  default = ""
}

target "default" {
  dockerfile = "./Dockerfile"
  context = "."
  target = "prod"
  tags = [
    strlen(GIT_TAG) > 0 ? "${IMAGE_NAME}:${GIT_TAG}": "",
    equal(GIT_TAG, LATEST_TAG) ? "${IMAGE_NAME}:latest": "",
    equal(GIT_TAG, "") && equal(BRANCH_NAME, "master") ? "${IMAGE_NAME}:edge": "",
    notequal(BRANCH_NAME, "master") && strlen(GIT_TAG) < 1 && strlen(GIT_SHA_TAG) > 0 ? "${IMAGE_NAME}:${GIT_SHA_TAG}": "",
  ]
  platforms = [
    "linux/amd64",
    strlen(GIT_TAG) > 0 ? "linux/arm64" : "",
  ]
}
