# ==== Compute Outputs

output "PublicInstanceIDs" {
  value = "${module.compute.server_id}"
}

output "PublicInstanceIPs" {
  value = "${module.compute.server_ip}"
}

output "PublicDns" {
  description = "List of public DNS names assigned to the instances. For EC2-VPC, this is only available if you've enabled DNS hostnames for your VPC"
  value       = "${module.compute.server_dns}"
}

output "BuildLink" {
  description = "Load Balancer DNS name"
  value       = "https://${module.networking.lb_dns}/builds"
}

output "CurlLink" {
  description = "Load Balancer DNS name"
  value       = "curl -d "jobs.json" -X POST https://${module.networking.lb_dns}/builds"
}
