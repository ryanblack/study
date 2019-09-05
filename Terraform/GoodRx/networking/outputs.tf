# ==== networking/outputs.tf ===
output "public_subnets" {
  value = "${aws_subnet.tf_public_subnet.*.id}"
}

output "public_sg" {
  value = "${aws_security_group.tf_public_sg.id}"
}

output "subnet_ips" {
  value = "${aws_subnet.tf_public_subnet.*.cidr_block}"
}

output "lb_dns" {
  value = "${aws_lb.tf_lb.dns_name}"
}

output "lb_target_group" {
  value = "${aws_lb_target_group.tf_lb_tg.arn}"
}