ALTER TABLE `Users`
  ADD COLUMN `user_m365_upn` varchar(255) DEFAULT NULL AFTER `user_email`,
  ADD COLUMN `user_m365_object_id` varchar(255) DEFAULT NULL AFTER `user_m365_upn`,
  ADD COLUMN `user_m365_tenant_id` varchar(255) DEFAULT NULL AFTER `user_m365_object_id`,
  ADD COLUMN `user_auth_provider` varchar(50) NOT NULL DEFAULT 'local' AFTER `user_password`,
  ADD COLUMN `user_last_login_utc` datetime DEFAULT NULL AFTER `user_auth_provider`;

CREATE UNIQUE INDEX `ux_users_m365_object_tenant`
  ON `Users` (`user_m365_object_id`, `user_m365_tenant_id`);

CREATE INDEX `ix_users_m365_upn`
  ON `Users` (`user_m365_upn`);

UPDATE `Users`
SET
  `user_m365_upn` = LOWER(`user_email`)
WHERE `user_m365_upn` IS NULL;
