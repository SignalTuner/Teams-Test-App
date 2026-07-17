-- New SignalTuner Teams meeting workflow.
-- This replaces the obsolete bot-populated TeamsMeetingParticipants roster cache.
-- Participants are authenticated SignalTuner users who opened the Teams app and joined the session.

ALTER TABLE `Users`
  ADD COLUMN `user_credits` int NOT NULL DEFAULT 5,
  ADD COLUMN `user_last_free_credit_grant_date` datetime DEFAULT NULL,
  ADD COLUMN `user_display_name` varchar(255) DEFAULT NULL;

CREATE TABLE IF NOT EXISTS `UserAuthIdentities` (
  `auth_identity_id` int NOT NULL AUTO_INCREMENT,
  `user_id` int NOT NULL,
  `provider` varchar(50) NOT NULL,
  `provider_subject_id` varchar(255) NOT NULL,
  `provider_tenant_id` varchar(255) DEFAULT NULL,
  `email_at_provider` varchar(255) DEFAULT NULL,
  `is_email_verified` tinyint(1) NOT NULL DEFAULT 0,
  `created_at` datetime NOT NULL DEFAULT CURRENT_TIMESTAMP,
  `updated_at` datetime NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (`auth_identity_id`),
  UNIQUE KEY `ux_user_auth_identity_provider_subject_tenant` (`provider`, `provider_subject_id`, `provider_tenant_id`),
  KEY `ix_user_auth_identities_user` (`user_id`),
  CONSTRAINT `fk_user_auth_identities_user`
    FOREIGN KEY (`user_id`) REFERENCES `Users` (`user_id`)
    ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS `TeamsMeetingSessions` (
  `meeting_session_id` int NOT NULL AUTO_INCREMENT,
  `teams_meeting_id` varchar(384) NOT NULL,
  `teams_conversation_id` varchar(384) DEFAULT NULL,
  `teams_tenant_id` varchar(255) DEFAULT NULL,
  `meeting_title` varchar(255) DEFAULT NULL,
  `created_at` datetime NOT NULL DEFAULT CURRENT_TIMESTAMP,
  `last_seen_at` datetime NOT NULL DEFAULT CURRENT_TIMESTAMP,
  PRIMARY KEY (`meeting_session_id`),
  KEY `ix_teams_meeting_sessions_meeting_id` (`teams_meeting_id`),
  KEY `ix_teams_meeting_sessions_tenant` (`teams_tenant_id`)
);

DROP TABLE IF EXISTS `TeamsMeetingParticipants`;

CREATE TABLE `TeamsMeetingParticipants` (
  `meeting_session_participant_id` int NOT NULL AUTO_INCREMENT,
  `meeting_session_id` int NOT NULL,
  `user_id` int NOT NULL,
  `joined_at` datetime NOT NULL DEFAULT CURRENT_TIMESTAMP,
  `last_seen_at` datetime NOT NULL DEFAULT CURRENT_TIMESTAMP,
  `display_name_snapshot` varchar(255) DEFAULT NULL,
  `auth_provider` varchar(50) DEFAULT NULL,
  PRIMARY KEY (`meeting_session_participant_id`),
  UNIQUE KEY `ux_teams_meeting_participants_session_user` (`meeting_session_id`, `user_id`),
  KEY `ix_teams_meeting_participants_user` (`user_id`),
  CONSTRAINT `fk_teams_meeting_participants_session`
    FOREIGN KEY (`meeting_session_id`) REFERENCES `TeamsMeetingSessions` (`meeting_session_id`)
    ON DELETE CASCADE,
  CONSTRAINT `fk_teams_meeting_participants_user`
    FOREIGN KEY (`user_id`) REFERENCES `Users` (`user_id`)
    ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS `CreditTransactions` (
  `credit_transaction_id` int NOT NULL AUTO_INCREMENT,
  `user_id` int NOT NULL,
  `amount` int NOT NULL,
  `reason` varchar(100) NOT NULL,
  `related_meeting_session_id` int DEFAULT NULL,
  `related_target_user_id` int DEFAULT NULL,
  `created_at` datetime NOT NULL DEFAULT CURRENT_TIMESTAMP,
  PRIMARY KEY (`credit_transaction_id`),
  KEY `ix_credit_transactions_user` (`user_id`),
  KEY `ix_credit_transactions_meeting_session` (`related_meeting_session_id`),
  CONSTRAINT `fk_credit_transactions_user`
    FOREIGN KEY (`user_id`) REFERENCES `Users` (`user_id`)
    ON DELETE CASCADE,
  CONSTRAINT `fk_credit_transactions_meeting_session`
    FOREIGN KEY (`related_meeting_session_id`) REFERENCES `TeamsMeetingSessions` (`meeting_session_id`)
    ON DELETE SET NULL,
  CONSTRAINT `fk_credit_transactions_target_user`
    FOREIGN KEY (`related_target_user_id`) REFERENCES `Users` (`user_id`)
    ON DELETE SET NULL
);

CREATE TABLE IF NOT EXISTS `EmailMagicCodes` (
  `email_magic_code_id` int NOT NULL AUTO_INCREMENT,
  `email` varchar(255) NOT NULL,
  `code_hash` varchar(255) NOT NULL,
  `expires_at` datetime NOT NULL,
  `used_at` datetime DEFAULT NULL,
  `created_at` datetime NOT NULL DEFAULT CURRENT_TIMESTAMP,
  `request_ip` varchar(64) DEFAULT NULL,
  `user_agent` varchar(512) DEFAULT NULL,
  PRIMARY KEY (`email_magic_code_id`),
  KEY `ix_email_magic_codes_email` (`email`),
  KEY `ix_email_magic_codes_expires` (`expires_at`)
);
