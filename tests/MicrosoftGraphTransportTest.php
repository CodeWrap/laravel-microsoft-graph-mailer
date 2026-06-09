<?php

use Codewrap\MicrosoftGraphMailer\Transport\MicrosoftGraphTransport;
use Illuminate\Http\Client\ConnectionException;
use Illuminate\Support\Facades\Cache;
use Illuminate\Support\Facades\Http;
use Symfony\Component\Mailer\Exception\TransportException;
use Symfony\Component\Mime\Email;

beforeEach(function () {
    Http::preventStrayRequests();
    Cache::flush();
});

function fakeTokenAndSend(): void
{
    Http::fake([
        'login.microsoftonline.com/*' => Http::response([
            'access_token' => 'fake-token',
            'expires_in' => 3600,
        ]),
        'graph.microsoft.com/*' => Http::response(null, 202),
    ]);
}

function buildEmail(): Email
{
    return (new Email())
        ->from('sender@example.com')
        ->to('recipient@example.com')
        ->subject('Test Subject')
        ->html('<p>Hello</p>');
}

function transport(array $config = []): MicrosoftGraphTransport
{
    return new MicrosoftGraphTransport(array_merge([
        'transport' => 'microsoft-graph',
        'tenant' => 'test-tenant',
        'client' => 'test-client',
        'secret' => 'test-secret',
        'save_to_sent_items' => true,
    ], $config));
}

function sendEmail(?Email $email = null, array $config = []): \Symfony\Component\Mailer\SentMessage
{
    return transport($config)->send($email ?? buildEmail());
}

function getSendMailRequest(): \Illuminate\Http\Client\Request
{
    $pairs = Http::recorded(fn ($r) => str_contains($r->url(), 'sendMail'));
    expect($pairs)->not->toBeEmpty();

    return $pairs->first()[0];
}

// --- Registration ---

test('transport resolves from mail manager', function () {
    $transport = app('mail.manager')
        ->mailer('microsoft-graph')
        ->getSymfonyTransport();

    expect($transport)->toBeInstanceOf(MicrosoftGraphTransport::class);
    expect((string) $transport)->toBe('microsoft-graph');
});

// --- Payload ---

test('sends email with correct payload', function () {
    fakeTokenAndSend();
    sendEmail();

    $request = getSendMailRequest();
    $body = $request->data();

    expect($request->method())->toBe('POST');
    expect($request->hasHeader('Authorization'))->toBeTrue();
    expect($request->header('Authorization')[0])->toBe('Bearer fake-token');
    expect($body['message']['subject'])->toBe('Test Subject');
    expect($body['message']['body']['contentType'])->toBe('HTML');
    expect($body['message']['body']['content'])->toContain('Hello');
    expect($body['message']['toRecipients'][0]['emailAddress']['address'])->toBe('recipient@example.com');
    expect($body['saveToSentItems'])->toBeTrue();
});

test('sends to correct graph api url with url-encoded sender', function () {
    fakeTokenAndSend();
    sendEmail();

    $request = getSendMailRequest();
    expect($request->url())->toContain('/users/sender%40example.com/sendMail');
});

test('maps cc and bcc recipients', function () {
    fakeTokenAndSend();

    $email = buildEmail();
    $email->cc('cc@example.com');
    $email->bcc('bcc@example.com');
    sendEmail($email);

    $body = getSendMailRequest()->data();

    expect($body['message']['ccRecipients'][0]['emailAddress']['address'])->toBe('cc@example.com');
    expect($body['message']['bccRecipients'][0]['emailAddress']['address'])->toBe('bcc@example.com');
});

test('maps reply-to with name', function () {
    fakeTokenAndSend();

    $email = buildEmail();
    $email->replyTo(new \Symfony\Component\Mime\Address('reply@example.com', 'Reply Name'));
    sendEmail($email);

    $replyTo = getSendMailRequest()->data()['message']['replyTo'][0];

    expect($replyTo['emailAddress']['address'])->toBe('reply@example.com');
    expect($replyTo['emailAddress']['name'])->toBe('Reply Name');
});

test('maps multiple to recipients', function () {
    fakeTokenAndSend();

    $email = buildEmail();
    $email->addTo('second@example.com');
    sendEmail($email);

    $recipients = getSendMailRequest()->data()['message']['toRecipients'];
    $addresses = array_column(array_column($recipients, 'emailAddress'), 'address');

    expect($addresses)->toContain('recipient@example.com');
    expect($addresses)->toContain('second@example.com');
});

test('omits optional fields when empty', function () {
    fakeTokenAndSend();
    sendEmail();

    $body = getSendMailRequest()->data();

    expect($body['message'])->not->toHaveKey('ccRecipients');
    expect($body['message'])->not->toHaveKey('bccRecipients');
    expect($body['message'])->not->toHaveKey('replyTo');
    expect($body['message'])->not->toHaveKey('attachments');
});

test('sends plain text when no html body', function () {
    fakeTokenAndSend();

    $email = (new Email())
        ->from('sender@example.com')
        ->to('recipient@example.com')
        ->subject('Plain')
        ->text('Just text');
    sendEmail($email);

    $body = getSendMailRequest()->data();

    expect($body['message']['body']['contentType'])->toBe('Text');
    expect($body['message']['body']['content'])->toBe('Just text');
});

test('html body takes precedence over text', function () {
    fakeTokenAndSend();

    $email = (new Email())
        ->from('sender@example.com')
        ->to('recipient@example.com')
        ->subject('Both')
        ->text('plain')
        ->html('<p>rich</p>');
    sendEmail($email);

    $body = getSendMailRequest()->data();

    expect($body['message']['body']['contentType'])->toBe('HTML');
    expect($body['message']['body']['content'])->toContain('rich');
});

test('save_to_sent_items false is passed', function () {
    fakeTokenAndSend();
    sendEmail(null, ['save_to_sent_items' => false]);

    expect(getSendMailRequest()->data()['saveToSentItems'])->toBeFalse();
});

// --- Custom Headers ---

test('passes only x- custom headers with values', function () {
    fakeTokenAndSend();

    $email = buildEmail();
    $email->getHeaders()->addTextHeader('X-Custom-Tag', 'my-tag');
    $email->getHeaders()->addTextHeader('X-Request-ID', 'abc-123');
    sendEmail($email);

    $headers = getSendMailRequest()->data()['message']['internetMessageHeaders'];
    $map = array_column($headers, 'value', 'name');

    expect($map)->toHaveKey('X-Custom-Tag', 'my-tag');
    expect($map)->toHaveKey('X-Request-ID', 'abc-123');

    foreach (array_keys($map) as $name) {
        expect(strtolower($name))->toStartWith('x-');
    }
});

test('excludes standard headers from internet message headers', function () {
    fakeTokenAndSend();
    sendEmail();

    $headers = getSendMailRequest()->data()['message']['internetMessageHeaders'] ?? [];
    $names = array_map(fn ($h) => strtolower($h['name']), $headers);

    expect($names)->not->toContain('from');
    expect($names)->not->toContain('to');
    expect($names)->not->toContain('subject');
    expect($names)->not->toContain('content-type');
    expect($names)->not->toContain('mime-version');
    expect($names)->not->toContain('date');
    expect($names)->not->toContain('message-id');
});

// --- Attachments ---

test('handles file attachment', function () {
    fakeTokenAndSend();

    $email = buildEmail();
    $email->attach('small content', 'test.txt', 'text/plain');
    sendEmail($email);

    $attachments = getSendMailRequest()->data()['message']['attachments'];

    expect($attachments)->toHaveCount(1);
    expect($attachments[0]['@odata.type'])->toBe('#microsoft.graph.fileAttachment');
    expect($attachments[0]['name'])->toBe('test.txt');
    expect($attachments[0]['contentType'])->toBe('text/plain');
    expect(base64_decode($attachments[0]['contentBytes']))->toBe('small content');
    expect($attachments[0])->not->toHaveKey('isInline');
});

test('handles inline attachment with content id', function () {
    fakeTokenAndSend();

    $email = buildEmail();
    $email->embed('image-data', 'logo.png', 'image/png');
    sendEmail($email);

    $attachments = getSendMailRequest()->data()['message']['attachments'];

    expect($attachments)->toHaveCount(1);
    expect($attachments[0]['isInline'])->toBeTrue();
    expect($attachments[0]['contentId'])->not->toBeEmpty();
    expect($attachments[0]['name'])->toBe('logo.png');
});

test('accepts attachment at exactly 3mb', function () {
    fakeTokenAndSend();

    $email = buildEmail();
    $email->attach(str_repeat('x', 3 * 1024 * 1024), 'exact.bin', 'application/octet-stream');
    sendEmail($email);

    $attachments = getSendMailRequest()->data()['message']['attachments'];
    expect($attachments)->toHaveCount(1);
});

test('rejects attachment over 3mb', function () {
    fakeTokenAndSend();

    $email = buildEmail();
    $email->attach(str_repeat('x', 3 * 1024 * 1024 + 1), 'big.bin', 'application/octet-stream');

    expect(fn () => sendEmail($email))
        ->toThrow(TransportException::class, '3 MB limit');

    Http::assertNotSent(fn ($r) => str_contains($r->url(), 'sendMail'));
});

// --- Sender / User ---

test('uses configured user for graph url', function () {
    fakeTokenAndSend();
    sendEmail(null, ['user' => 'mailer@contras.net']);

    expect(getSendMailRequest()->url())->toContain('/users/mailer%40contras.net/sendMail');
});

test('configured user takes precedence over from address', function () {
    fakeTokenAndSend();

    $email = (new Email())
        ->from('different@example.com')
        ->to('recipient@example.com')
        ->subject('Test')
        ->html('<p>test</p>');

    sendEmail($email, ['user' => 'fixed@contras.net']);

    expect(getSendMailRequest()->url())->toContain('/users/fixed%40contras.net/sendMail');
});

// --- Success Headers ---

test('sets message id headers after successful send', function () {
    fakeTokenAndSend();
    $sent = sendEmail();

    $headers = $sent->getOriginalMessage()->getHeaders();

    expect($headers->has('X-Message-ID'))->toBeTrue();
    expect($headers->has('X-Graph-Message-ID'))->toBeTrue();

    $messageId = $headers->get('X-Message-ID')->getBodyAsString();
    expect($messageId)->toContain('@microsoft-graph');
    expect($headers->get('X-Graph-Message-ID')->getBodyAsString())->toBe($messageId);
});

// --- Token ---

test('token request uses correct credentials and format', function () {
    fakeTokenAndSend();
    sendEmail();

    $tokenRequests = Http::recorded(fn ($r) => str_contains($r->url(), 'oauth2/v2.0/token'));
    expect($tokenRequests)->toHaveCount(1);

    $request = $tokenRequests->first()[0];

    expect($request->method())->toBe('POST');
    expect($request->url())->toContain('test-tenant');
    expect($request->data()['client_id'])->toBe('test-client');
    expect($request->data()['client_secret'])->toBe('test-secret');
    expect($request->data()['scope'])->toBe('https://graph.microsoft.com/.default');
    expect($request->data()['grant_type'])->toBe('client_credentials');
});

test('caches access token across sends', function () {
    fakeTokenAndSend();

    sendEmail();
    sendEmail();

    $tokenRequests = Http::recorded(fn ($r) => str_contains($r->url(), 'oauth2/v2.0/token'));
    expect($tokenRequests)->toHaveCount(1);
});

// --- 401 Retry ---

test('retries with fresh token on 401', function () {
    Http::fake([
        'login.microsoftonline.com/*' => Http::sequence()
            ->push(['access_token' => 'fresh-token', 'expires_in' => 3600])
            ->push(['access_token' => 'fresh-token-2', 'expires_in' => 3600]),
        'graph.microsoft.com/*' => Http::sequence()
            ->push(null, 401)
            ->push(null, 202),
    ]);

    Cache::put('msgraph-mailer:test-tenant:test-client', 'stale-token', 3600);

    sendEmail();

    $sendRequests = Http::recorded(fn ($r) => str_contains($r->url(), 'sendMail'));
    expect($sendRequests)->toHaveCount(2);

    $all = $sendRequests->values();
    $firstAuth = $all[0][0]->header('Authorization')[0];
    $secondAuth = $all[1][0]->header('Authorization')[0];

    expect($firstAuth)->toBe('Bearer stale-token');
    expect($secondAuth)->not->toBe('Bearer stale-token');
});

test('throws after second 401 without infinite retry', function () {
    Http::fake([
        'login.microsoftonline.com/*' => Http::response([
            'access_token' => 'still-bad',
            'expires_in' => 3600,
        ]),
        'graph.microsoft.com/*' => Http::response(null, 401),
    ]);

    expect(fn () => sendEmail())
        ->toThrow(TransportException::class);

    $sendRequests = Http::recorded(fn ($r) => str_contains($r->url(), 'sendMail'));
    expect($sendRequests)->toHaveCount(2);
});

// --- Error Handling ---

test('throws on 403 forbidden', function () {
    Http::fake([
        'login.microsoftonline.com/*' => Http::response([
            'access_token' => 'fake-token',
            'expires_in' => 3600,
        ]),
        'graph.microsoft.com/*' => Http::response([
            'error' => ['code' => 'ErrorAccessDenied', 'message' => 'Access is denied.'],
        ], 403),
    ]);

    expect(fn () => sendEmail())
        ->toThrow(TransportException::class, 'Access is denied');
});

test('throws on 500 server error', function () {
    Http::fake([
        'login.microsoftonline.com/*' => Http::response([
            'access_token' => 'fake-token',
            'expires_in' => 3600,
        ]),
        'graph.microsoft.com/*' => Http::response([
            'error' => ['code' => 'InternalServerError', 'message' => 'An internal error occurred.'],
        ], 500),
    ]);

    expect(fn () => sendEmail())
        ->toThrow(TransportException::class, 'internal error');
});

test('error message does not contain secrets', function () {
    Http::fake([
        'login.microsoftonline.com/*' => Http::response([
            'access_token' => 'fake-token',
            'expires_in' => 3600,
        ]),
        'graph.microsoft.com/*' => Http::response([
            'error' => ['code' => 'Forbidden', 'message' => 'Denied'],
        ], 403),
    ]);

    try {
        sendEmail();
    } catch (TransportException $e) {
        expect($e->getMessage())->not->toContain('test-secret');
        expect($e->getMessage())->not->toContain('fake-token');
        expect($e->getMessage())->toContain('Denied');

        return;
    }

    test()->fail('Expected TransportException was not thrown.');
});

test('throws on token fetch failure', function () {
    Http::fake([
        'login.microsoftonline.com/*' => Http::response(['error' => 'invalid_client'], 401),
    ]);

    expect(fn () => sendEmail())
        ->toThrow(TransportException::class, 'access token');
});

test('throws TransportException on token fetch connection error', function () {
    Http::fake([
        'login.microsoftonline.com/*' => fn () => throw new ConnectionException('Connection refused'),
    ]);

    expect(fn () => sendEmail())
        ->toThrow(TransportException::class, 'Connection refused');
});

test('throws TransportException on send connection error', function () {
    Http::fake([
        'login.microsoftonline.com/*' => Http::response([
            'access_token' => 'fake-token',
            'expires_in' => 3600,
        ]),
        'graph.microsoft.com/*' => fn () => throw new ConnectionException('Network unreachable'),
    ]);

    expect(fn () => sendEmail())
        ->toThrow(TransportException::class, 'Network unreachable');
});
