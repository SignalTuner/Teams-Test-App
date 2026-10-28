import React from "react";
import TemporaryActivationCode from "./TemporaryActivationCode";
import { SignalTunerLogo } from "./DashboardHeader";
import { useExternalCommerce } from "./TeamsHostProvider";
import downloadStepIcon from "./assets/download.svg";
import personCheckIcon from "./assets/person-check.svg";
import refreshStepIcon from "./assets/refresh.svg";
import clientPreview from "./assets/client.png";
const CLIENT_PROMPT_REFRESH_INTERVAL_MS = 20000;
const CLIENT_PROMPT_COPIED_REFRESH_INTERVAL_MS = 2000;
const CLIENT_PROMPT_COPIED_REFRESH_DURATION_MS = 60000;
function getDownloadUrl(): string | null {
  const userAgent = window.navigator.userAgent;

  if (/Windows NT/i.test(userAgent)) {
    return "https://signaltuner.com/update/SignalTuner.exe";
  }

  if (/Mac OS|Macintosh|MacIntel/i.test(userAgent)) {
    return "https://signaltuner.com/update/macos/SignalTuner.dmg";
  }

  return null;
}

export default function ClientPrompt({
  activationCodeError,
  isLoading,
  onContinue,
  onRefresh,
  onSignOut,
}: {
  activationCodeError: string | null;
  isLoading: boolean;
  onContinue: () => void;
  onRefresh: () => Promise<void>;
  onSignOut: () => void;
}) {
  const allowCommerce = useExternalCommerce();
  // Host context controls visibility; user agent only selects a desktop installer.
  const downloadUrl = allowCommerce ? getDownloadUrl() : null;
  const browserLaunchUrl = "https://www.signaltuner.com/Launch.html";

  const [rapidPollingEndsAt, setRapidPollingEndsAt] = React.useState<number | null>(null);
  const onRefreshRef = React.useRef(onRefresh);
  const refreshInFlightRef = React.useRef(false);

  React.useEffect(() => {
    onRefreshRef.current = onRefresh;
  }, [onRefresh]);

  const runRefresh = React.useCallback(async () => {
    if (refreshInFlightRef.current) {
      return;
    }

    refreshInFlightRef.current = true;

    try {
      await onRefreshRef.current();
    } finally {
      refreshInFlightRef.current = false;
    }
  }, []);

  const handleActivationCodeCopied = React.useCallback(() => {
    setRapidPollingEndsAt(Date.now() + CLIENT_PROMPT_COPIED_REFRESH_DURATION_MS);
    void runRefresh();
  }, [runRefresh]);

  const refreshIntervalMs =
    rapidPollingEndsAt === null ? CLIENT_PROMPT_REFRESH_INTERVAL_MS : CLIENT_PROMPT_COPIED_REFRESH_INTERVAL_MS;

  React.useEffect(() => {
    const intervalId = window.setInterval(() => {
      void runRefresh();
    }, refreshIntervalMs);

    return () => window.clearInterval(intervalId);
  }, [refreshIntervalMs, runRefresh]);

  React.useEffect(() => {
    if (rapidPollingEndsAt === null) {
      return;
    }

    const timeoutId = window.setTimeout(
      () => setRapidPollingEndsAt(null),
      Math.max(rapidPollingEndsAt - Date.now(), 0)
    );

    return () => window.clearTimeout(timeoutId);
  }, [rapidPollingEndsAt]);

  const setupSteps = [
    { icon: downloadStepIcon, isImage: true, title: "Set up the SignalTuner desktop app on your computer", description: allowCommerce ? "This downloads a desktop installer outside Microsoft Teams. ↗" : "Use your computer to connect the desktop app." },
    {
      icon: personCheckIcon,
      isImage: true,
      title: "Open the app and enter your activation code, or sign in with email and password",
      description: allowCommerce ? "The launch link opens outside Microsoft Teams. ↗ This links the app to your account." : "This links the app to your account.",
    },
    { icon: refreshStepIcon, isImage: true, isFlipped: true, title: "Return here and refresh your status", description: "We'll confirm when your device is connected." },
  ];

  const benefits = [
    "Live network telemetry during meetings",
    "Device insights for better visibility",
    "Better troubleshooting and faster resolutions",
  ];

  return (
    <section className="clientPrompt">
      <div className="clientPromptIntro">
        <SignalTunerLogo className="clientPromptLogo" />
        <h1>Connect the desktop client</h1>
        <p className="clientPromptLead">
          Connect the SignalTuner desktop app on your computer to share your live Wi-Fi, network, and device telemetry
          during meetings.
        </p>

        <div className="setupSteps">
          {setupSteps.map((step, index) => (
            <div className="setupStep" key={step.title}>
              <div className="setupStepRail" aria-hidden="true">
                <span className="setupStepNumber">{index + 1}</span>
                {index < setupSteps.length - 1 && <span className="setupStepConnector" />}
              </div>
              <div className="setupStepIcon">
                {index === 0 && downloadUrl ? (
                  <a className="setupStepIconButton" href={downloadUrl} target="_blank" rel="noopener noreferrer" aria-label="Download SignalTuner outside Microsoft Teams">
                    <img src={step.icon} alt="" />
                  </a>
                ) : index === 1 && allowCommerce ? (
                  <a
                    className="setupStepIconButton"
                    href={browserLaunchUrl}
                    target="_blank"
                    rel="noopener noreferrer"
                    aria-label="Open SignalTuner desktop client outside Microsoft Teams"
                  >
                    <img src={step.icon} alt="" />
                  </a>
                ) : index === 2 ? (
                  <button
                    className="setupStepIconButton"
                    disabled={isLoading}
                    onClick={() => void runRefresh()}
                    type="button"
                    aria-label="Refresh status"
                  >
                    {isLoading ? <Spinner /> : <img className="setupStepIconFlipped" src={step.icon} alt="" />}
                  </button>
                ) : step.isImage ? (
                  <img src={step.icon} alt="" />
                ) : step.icon}
              </div>
              <div className="setupStepCopy">
                <strong>{step.title}</strong>
                <p>{step.description}</p>
              </div>
            </div>
          ))}
        </div>

        <div className="clientBenefits">
          <h2>Why connect it?</h2>
          {benefits.map((benefit) => (
            <div className="benefitItem" key={benefit}>
              <span className="benefitCheck" aria-hidden="true">✓</span>
              <span>{benefit}</span>
            </div>
          ))}
        </div>

        <div className="clientInfoNote">
          <span className="infoGlyph" aria-hidden="true">i</span>
          <span>You can continue without device data, but meeting insights will be limited until your desktop client is connected.</span>
        </div>
      </div>

      <div className="activationPanel">
        <div className="activationPanelHeader">
          <span className="eyebrow">GET STARTED</span>
          <h2>Your activation code</h2>
          <p>Use this code in the desktop client.</p>
        </div>
        <TemporaryActivationCode onCopied={handleActivationCodeCopied} />
        {activationCodeError && <p className="activationError">{activationCodeError}</p>}

        <div className="desktopPreview" aria-label="Illustration of where to enter the activation code">
          <img className="desktopPreviewImage" src={clientPreview} alt="SignalTuner desktop client sign-in screen" />
        </div>

        <div className="clientActions">
          {downloadUrl ? (
            <a className="primaryButton buttonLink clientDownloadButton" href={downloadUrl} target="_blank" rel="noopener noreferrer">
              <img className="actionIcon actionIconImage" src={downloadStepIcon} alt="" aria-hidden="true" /> Download outside Teams ↗
            </a>
          ) : (
            <span className="inlineNote">Connect the desktop app on your computer, then refresh status here.</span>
          )}
          <button className="secondaryButton clientActionButton" disabled={isLoading} onClick={() => void runRefresh()} type="button">
            {isLoading ? <Spinner /> : <span className="actionIcon" aria-hidden="true">↻</span>} Refresh status
          </button>
          <button className="secondaryButton clientActionButton clientTertiaryButton" disabled={isLoading} onClick={onContinue} type="button">
            Continue without syncing
          </button>
        </div>
        <button className="clientSignOut" disabled={isLoading} onClick={onSignOut} type="button">Sign out</button>
      </div>
    </section>
  );
}


function Spinner() { return <span className="spinner" aria-hidden="true" />; }
