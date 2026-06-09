<?php

namespace Codewrap\MicrosoftGraphMailer\Tests;

use Codewrap\MicrosoftGraphMailer\MicrosoftGraphMailerServiceProvider;
use Orchestra\Testbench\TestCase as Orchestra;

class TestCase extends Orchestra
{
    protected function getPackageProviders($app): array
    {
        return [
            MicrosoftGraphMailerServiceProvider::class,
        ];
    }

    protected function defineEnvironment($app): void
    {
        $app['config']->set('mail.mailers.microsoft-graph', [
            'transport' => 'microsoft-graph',
            'tenant' => 'test-tenant',
            'client' => 'test-client',
            'secret' => 'test-secret',
            'save_to_sent_items' => true,
        ]);

        $app['config']->set('mail.from.address', 'sender@example.com');
        $app['config']->set('mail.from.name', 'Test Sender');
    }
}
