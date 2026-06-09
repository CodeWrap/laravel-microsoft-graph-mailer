# Laravel Microsoft Graph Mailer

A Laravel mail driver that sends email via the Microsoft Graph API using OAuth2 client credentials. Zero external dependencies beyond Laravel itself.

## Requirements

- PHP 8.2+
- Laravel 11, 12, or 13
- An Azure AD / Entra ID app registration with `Mail.Send` application permission

## Installation

```bash
composer require codewrap/laravel-microsoft-graph-mailer
```

## Azure AD Setup

1. Go to **Azure Portal** > **App registrations** > **New registration**
2. Name your app (e.g. "My App Mailer"), select **Single tenant**
3. Go to **API permissions** > **Add permission** > **Microsoft Graph** > **Application permissions** > **Mail.Send** > **Grant admin consent**
4. Go to **Certificates & secrets** > **New client secret** — copy the secret value
5. From the **Overview** page, copy the **Application (client) ID** and **Directory (tenant) ID**

### Restrict Sender Addresses (Recommended)

By default, `Mail.Send` allows sending as any mailbox in your tenant. To restrict which mailboxes the app can send from, configure an **Application Access Policy** in Exchange Online:

```powershell
New-ApplicationAccessPolicy -AppId <client-id> -PolicyScopeGroupId <mail-enabled-security-group> -AccessRight RestrictAccess
```

## Configuration

Add the mailer to `config/mail.php`:

```php
'mailers' => [
    'microsoft-graph' => [
        'transport' => 'microsoft-graph',
        'tenant' => env('MAIL_MSGRAPH_TENANT'),
        'client' => env('MAIL_MSGRAPH_CLIENT'),
        'secret' => env('MAIL_MSGRAPH_SECRET'),
        'save_to_sent_items' => env('MAIL_MSGRAPH_SAVE_TO_SENT', true),
    ],
],
```

Add to your `.env`:

```
MAIL_MAILER=microsoft-graph
MAIL_MSGRAPH_TENANT=your-tenant-id
MAIL_MSGRAPH_CLIENT=your-client-id
MAIL_MSGRAPH_SECRET=your-client-secret
MAIL_FROM_ADDRESS=noreply@yourdomain.com
MAIL_FROM_NAME="Your App"
```

The `MAIL_FROM_ADDRESS` must be a valid mailbox in your Microsoft 365 tenant.

## Usage

Once configured as the default mailer, all `Mail::send()` / `Mail::to()` calls use Microsoft Graph automatically. No code changes needed.

## Limitations

- **3 MB attachment limit** — Direct file attachments in `sendMail` are limited to 3 MB. Attachments between 3-150 MB require a draft + upload session flow (not yet supported).
- **`X-Message-ID`** — Graph's `sendMail` returns no message ID. The driver generates a local correlation UUID, not an Exchange message ID.

## National Clouds

For national cloud deployments (US Gov, China), override the API endpoints in your `.env`:

```
MSGRAPH_API_BASE_URL=https://graph.microsoft.us/v1.0
MSGRAPH_AUTH_URL=https://login.microsoftonline.us
```

## License

MIT
