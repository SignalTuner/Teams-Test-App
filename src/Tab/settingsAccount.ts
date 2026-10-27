export type ReportAccount = {
  purchasedCredits: number;
  availableReports: number;
  teamAllocatedCredits?: number;
  teamFundAvailable?: number;
  isCustomer?: boolean;
  subscription: null | {
    name: string; status: string; remaining: number; allowance: number;
    renewalDate: string | null; cancelAtPeriodEnd: boolean; paidPeriod: boolean;
  };
};

export async function loadReportAccount(apiBaseUrl: string, token: string, signal: AbortSignal): Promise<ReportAccount> {
  const response = await fetch(`${apiBaseUrl}/api/credits/account`, {
    headers: { Authorization: `Bearer ${token}` }, signal,
  });
  if (!response.ok) throw new Error(response.status === 401
    ? "Your session has expired. Sign in again to view your account."
    : "Your account could not be loaded. Please try again.");
  const data = await response.json();
  if (!Number.isInteger(data?.purchasedCredits) || !Number.isInteger(data?.availableReports)) {
    throw new Error("Your account could not be loaded. Please try again.");
  }
  return data;
}

export async function saveAccountProfile(apiBaseUrl: string, token: string, userId: number,
  firstName: string, lastName: string, email: string): Promise<void> {
  const form = new FormData();
  form.append("userId", String(userId));
  form.append("userFirstName", firstName.trim());
  form.append("userLastName", lastName.trim());
  form.append("userEmail", email.trim());
  const response = await fetch(`${apiBaseUrl}/api/User/UpdateUser`, {
    method: "PUT", headers: { Authorization: `Bearer ${token}` }, body: form,
  });
  const result = await response.json().catch(() => null);
  if (response.ok && result?.status === "success") return;
  // Do not render arbitrary server/exception text in the account form.
  const messages = new Set([
    "Enter a valid email address.", "First name and last name are required.",
    "First name and last name must be 255 characters or fewer.",
    "Create a password before updating your email.", "That email is already in use.",
  ]);
  throw new Error(response.status === 401 ? "Your session has expired. Sign in again."
    : response.status < 500 && messages.has(result?.message) ? result.message
    : "Your profile could not be saved. Please try again.");
}

export function subscriptionStatus(subscription: ReportAccount["subscription"]) {
  if (!subscription) return "Pay as you go";
  if (subscription.status === "active") return subscription.cancelAtPeriodEnd ? "Ends after this period" : "Active";
  const labels: Record<string, string> = { canceled: "Canceled", cancelled: "Canceled", past_due: "Payment overdue",
    unpaid: "Payment required", incomplete: "Payment pending", incomplete_expired: "Payment expired",
    trialing: "Trial", paused: "Paused", unsupported: "Needs attention" };
  return labels[subscription.status] || "Needs attention";
}

export function periodDate(value: string | null | undefined) {
  const date = value ? new Date(value) : null;
  return date && !Number.isNaN(date.getTime())
    ? date.toLocaleDateString(undefined, { month: "short", day: "numeric", year: "numeric" }) : null;
}
