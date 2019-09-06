aws_credentials = "/mnt/d/DISK/credentials"
aws_region = "us-west-2"
project_name = "GR-terraform"
vpc_cidr = "10.0.0.0/16"
public_cidrs = [
  "10.0.1.0/24",
  "10.0.2.0/24"
  ]
accessip = "0.0.0.0/0"
key_name = "tf_key"
public_key_path = "./id_rsa.pub"
server_instance_type = "t2.micro"
instance_count = 1
