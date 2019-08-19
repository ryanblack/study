provider "aws" {
  region                  = "${var.aws_region}"
  shared_credentials_file = "./credentials"
  profile                 = "terraform"
}

# Deploy Storage Resource

module "storage" {
  source = "./storage"
  project_name = "${var.project_name}"
}
