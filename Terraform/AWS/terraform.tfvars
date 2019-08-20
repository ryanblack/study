aws_region = "us-west-2"
project_name = "la-terraform"
vpc_cidr = "192.168.0.0/16"
public_cidrs = [
  "192.168.1.0/24",
  "192.168.2.0/24"
  ]
accessip = "0.0.0.0/0"
key_name = "tf_key"
public_key_path = "./id_rsa.pub"
server_instance_type = "t2.micro"
instance_count = 2
