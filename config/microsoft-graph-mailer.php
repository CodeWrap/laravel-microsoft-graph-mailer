<?php

return [

    'api_base_url' => env('MSGRAPH_API_BASE_URL', 'https://graph.microsoft.com/v1.0'),

    'auth_url' => env('MSGRAPH_AUTH_URL', 'https://login.microsoftonline.com'),

    /*
     * Seconds subtracted from the token's expires_in to refresh early.
     */
    'token_cache_skew' => 300,

];
