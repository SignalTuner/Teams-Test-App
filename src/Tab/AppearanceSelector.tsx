import "./AppearanceSelector.css";

export default function AppearanceSelector({ value, onChange }: {
  value: "light" | "dark";
  onChange: (value: "light" | "dark") => void;
}) {
  return <div className="appearanceSelector" role="group" aria-label="SignalTuner appearance">
    {(["light", "dark"] as const).map(option => <button key={option} type="button"
      aria-pressed={value === option} onClick={() => onChange(option)}>
      {option === "light" ? "Light" : "Dark"}
    </button>)}
  </div>;
}
