ssh-keygen -t rsa -N "" -f id_rsa && cp ~/.ssh/id_rsa* ./
terraform apply -auto-approve
