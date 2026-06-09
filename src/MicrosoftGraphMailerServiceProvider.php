<?php

namespace Codewrap\MicrosoftGraphMailer;

use Codewrap\MicrosoftGraphMailer\Transport\MicrosoftGraphTransport;
use Spatie\LaravelPackageTools\Package;
use Spatie\LaravelPackageTools\PackageServiceProvider;

class MicrosoftGraphMailerServiceProvider extends PackageServiceProvider
{
    public function configurePackage(Package $package): void
    {
        $package
            ->name('laravel-microsoft-graph-mailer')
            ->hasConfigFile();
    }

    public function packageBooted(): void
    {
        $this->app->afterResolving('mail.manager', function ($manager) {
            $manager->extend('microsoft-graph', function (array $config) {
                return new MicrosoftGraphTransport($config);
            });
        });
    }
}
