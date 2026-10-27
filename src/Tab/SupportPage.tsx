import React, { useEffect, useRef, useState } from "react";
import AppearanceSelector from "./AppearanceSelector";
import "./Support.css";

const SUPPORT_EMAIL = "support@signaltuner.com";

function EnvelopeIcon() {
  return <svg width="1em" height="1em" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" aria-hidden="true"><rect x="2" y="4" width="20" height="16" rx="2" /><path d="m3 6 9 7 9-7" /></svg>;
}

export default function SupportPage({ onSendSupportEmail, defaultEmail = "", defaultName = "", themePreference, onThemePreferenceChange }: {
  onSendSupportEmail: (request: { sender: string; subject: string; body: string }) => Promise<string>;
  defaultEmail?: string | null; defaultName?: string | null;
  themePreference: "light" | "dark";
  onThemePreferenceChange: (value: "light" | "dark") => void;
}) {
  const [fullName, setFullName] = useState("");
  const [email, setEmail] = useState("");
  const [subject, setSubject] = useState("");
  const [message, setMessage] = useState("");
  const [submitting, setSubmitting] = useState(false);
  const [feedback, setFeedback] = useState<{ type: "success" | "error"; text: string } | null>(null);
  const submittingRef = useRef(false);

  useEffect(() => {
    setFullName(defaultName || ""); setEmail(defaultEmail || "");
  }, [defaultName, defaultEmail]);

  const handleSubmit = async (event: React.FormEvent) => {
    event.preventDefault();
    if (submittingRef.current) return;

    const name = fullName.trim();
    const sender = email.trim();
    const topic = subject.trim();
    const body = message.trim();
    if (!name || !sender || !topic || !body) {
      setFeedback({ type: "error", text: "Please complete all fields." });
      return;
    }

    // The existing API has no name field; include it in the delivered message.
    const deliveredBody = `Name: ${name}\n\n${body}`;
    if (deliveredBody.length > 1000) {
      setFeedback({ type: "error", text: "Please shorten your message and try again." });
      return;
    }

    submittingRef.current = true;
    setSubmitting(true);
    setFeedback(null);
    try {
      await onSendSupportEmail({ sender, subject: topic, body: deliveredBody });
      setFullName("");
      setEmail("");
      setSubject("");
      setMessage("");
      setFeedback({ type: "success", text: "Your message was sent. We'll reply to the email address you provided." });
    } catch {
      setFeedback({ type: "error", text: "We couldn't send your message. Please try again." });
    } finally {
      submittingRef.current = false;
      setSubmitting(false);
    }
  };

  return <section className="support-page" aria-labelledby="support-heading">
    <div className="support-inner">
      <header className="support-header">
        <div className="support-header-copy">
          <p className="support-eyebrow">SIGNALTUNER SUPPORT</p>
          <h1 id="support-heading">Contact support</h1>
          <p>Need help with your account, downloads, or connectivity reports? Send us a message and our team will get back to you.</p>
        </div>
        <AppearanceSelector value={themePreference} onChange={onThemePreferenceChange} />
      </header>

      <div className="support-content">
        <section className="support-card" aria-labelledby="support-form-heading">
          <div className="support-card-heading">
            <span className="support-icon" aria-hidden="true"><EnvelopeIcon /></span>
            <div>
              <h2 id="support-form-heading">Send us a message</h2>
              <p>Fill out the form below and we'll get back to you as soon as possible.</p>
            </div>
          </div>

          <form className="support-form" aria-busy={submitting} onSubmit={handleSubmit}>
            <label htmlFor="support-name">Full name</label>
            <input id="support-name" type="text" autoComplete="name" placeholder="Your full name" required maxLength={80}
              value={fullName} onChange={event => setFullName(event.target.value)} />

            <label htmlFor="support-email">Email address</label>
            <input id="support-email" type="email" autoComplete="email" placeholder="you@company.com" required maxLength={50}
              value={email} onChange={event => setEmail(event.target.value)} />

            <label htmlFor="support-subject">Subject</label>
            <input id="support-subject" type="text" placeholder="What can we help you with?" required maxLength={80}
              value={subject} onChange={event => setSubject(event.target.value)} />

            <label htmlFor="support-message">Message</label>
            <textarea id="support-message" placeholder="Tell us more..." required maxLength={1000}
              value={message} onChange={event => setMessage(event.target.value)} />

            <button type="submit" disabled={submitting}>{submitting ? "Sending..." : "Send message"}</button>
            <p className="support-reply-note">We'll respond to the email address you provided.</p>
            {feedback && <p className={`support-feedback support-feedback--${feedback.type}`} role={feedback.type === "error" ? "alert" : "status"}>{feedback.text}</p>}
          </form>
        </section>

        <aside className="support-info" aria-label="Support contact information">
          <span className="support-info-icon" aria-hidden="true"><EnvelopeIcon /></span>
          <div><h2>Email support</h2><a href={`mailto:${SUPPORT_EMAIL}`}>{SUPPORT_EMAIL}</a></div>
        </aside>
      </div>
    </div>
  </section>;
}
