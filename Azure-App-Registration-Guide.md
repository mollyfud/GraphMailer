# Azure App Registration Guide for GraphMailer

This guide explains, step by step, how to create and configure the Azure application required by the GraphMailer library.

## Goal

Create an Azure app registration that can authenticate with Microsoft Graph using **application permissions** and send email with `Mail.Send`.

---

## Prerequisites

- An Azure tenant (Microsoft Entra ID)
- Access to the [Azure portal](https://portal.azure.com)
- Permission to register applications in the tenant
- A tenant administrator available to grant admin consent

---

## Step 1: Open App Registrations

1. Sign in to [https://portal.azure.com](https://portal.azure.com).
2. Go to **Microsoft Entra ID**.
3. In the left menu, select **App registrations**.
4. Select **New registration**.

---

## Step 2: Register the application

1. In **Name**, enter a clear name (example: `GraphMailer-Production`).
2. In **Supported account types**, select:
   - **Accounts in this organizational directory only** (single tenant), unless your scenario requires multi-tenant.
3. Leave **Redirect URI** empty (not required for this library’s app-only flow).
4. Select **Register**.

After creation, stay on the **Overview** page.

---

## Step 3: Copy Tenant ID and Client ID

From the app’s **Overview** page, copy and securely store:

- **Directory (tenant) ID** → use as `TenantId`
- **Application (client) ID** → use as `ClientId`

You will use these values in `AuthenticationData`.

---

## Step 4: Create a client secret

1. In the app menu, select **Certificates & secrets**.
2. In **Client secrets**, select **New client secret**.
3. Enter a description (example: `GraphMailer Secret`).
4. Choose an expiration period that matches your policy.
5. Select **Add**.
6. Immediately copy the **Value** of the secret and store it securely.

Use this secret value as `ClientSecret`.

> Important: You can only copy the secret value once. If lost, create a new secret.

---

## Step 5: Add Microsoft Graph API permissions

1. In the app menu, select **API permissions**.
2. Select **Add a permission**.
3. Select **Microsoft Graph**.
4. Select **Application permissions**.
5. Search for and add:
   - `Mail.Send` (required)
6. If you plan to send large attachments (over 3 MB), also add:
   - `Mail.ReadWrite` (required by large-file draft/upload flow)

---

## Step 6: Grant admin consent

1. On **API permissions**, confirm the permissions are listed under **Application**.
2. Select **Grant admin consent for <TenantName>**.
3. Confirm the action.
4. Verify the status changes to **Granted for <TenantName>**.

Without admin consent, token acquisition may succeed but Graph operations will fail with authorization errors.

---

## Step 7: (Recommended) Restrict mailbox scope in Exchange Online

By default, `Mail.Send` application permission can allow broad send capability in the tenant.
For least privilege, configure an **Exchange Online Application Access Policy** (or equivalent current control) to limit which mailboxes this app can access.

This is optional for basic setup but strongly recommended for production.

---

## Step 8: Configure GraphMailer

Use the three Azure values in your application:

- `TenantId` = Directory (tenant) ID
- `ClientId` = Application (client) ID
- `ClientSecret` = Client secret value

Example:

- `TenantId`: `xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx`
- `ClientId`: `yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy`
- `ClientSecret`: `<secret value>`

---

## Step 9: Validate with token diagnostics (recommended)

If you receive `401 Unauthorized` or `403 Forbidden`:

1. Use GraphMailer token dump functionality (`GetTokenDump` / `GetTokenDumpAsync`).
2. Open the token in [https://jwt.ms](https://jwt.ms).
3. Confirm `roles` includes:
   - `Mail.Send`
   - `Mail.ReadWrite` (if using large attachments)

If roles are missing:
- Recheck API permissions
- Re-grant admin consent
- Acquire a fresh token again

---

## Common setup mistakes

- Using **Delegated** permission instead of **Application** permission
- Forgetting to grant admin consent
- Copying the wrong secret field (must be **Value**, not Secret ID)
- Using incorrect tenant (wrong `TenantId`)
- Secret expired
- `fromAddress` mailbox not valid/licensed or blocked by org policy

---

## Rotation and operational recommendations

- Store secrets in a secure vault (for example, Azure Key Vault)
- Rotate client secrets before expiration
- Use separate app registrations per environment (Dev/Test/Prod)
- Review app permissions regularly

---

## Required values checklist

Before using GraphMailer, confirm you have:

- [ ] Tenant ID
- [ ] Client ID
- [ ] Client Secret (value)
- [ ] `Mail.Send` application permission
- [ ] Admin consent granted
- [ ] `Mail.ReadWrite` granted (only if sending attachments larger than 3 MB)
