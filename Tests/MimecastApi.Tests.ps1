Describe 'MimecastApi Module Tests' {
    BeforeAll {
        $ModulePath = Split-Path -Parent $PSScriptRoot
        Import-Module $ModulePath -Force
    }

    Context 'Module Loading' {
        It 'Imports successfully' {
            Get-Module MimecastApi | Should -Not -BeNull
        }

        It 'Has required functions exported' {
            $RequiredFunctions = @(
                'Connect-MimecastApi',
                'Disconnect-MimecastApi',
                'Set-MimecastApiConfiguration',
                'Get-MimecastApiConfiguration',
                'Get-MimecastApiSystemInfo'
            )
            
            $ExportedFunctions = Get-Command -Module MimecastApi
            foreach ($Function in $RequiredFunctions) {
                $ExportedFunctions.Name | Should -Contain $Function
            }
        }
    }

    Context 'Configuration' {
        It 'Sets and gets configuration' {
            Set-MimecastApiConfiguration -LogFilePath 'TestLog.log' -ApiVaultName 'TestVault'
            $config = Get-MimecastApiConfiguration
            $config.LogFilePath | Should -Be 'TestLog.log'
            $config.ApiVaultName | Should -Be 'TestVault'
        }
    }

    Context 'Module Initialization' {
        BeforeAll {
            # Mock Get-SecretVault
            Mock Get-SecretVault { $null }
            Mock Register-SecretVault { $true }
            
            # Mock Set-ApiSecret
            Mock Set-ApiSecret { $true }
            
            # Mock Get-ApiSecret for checking existing secrets
            Mock Get-ApiSecret { $null }
            
            # Mock Read-Host
            Mock Read-Host { ConvertTo-SecureString "test-secret" -AsPlainText -Force }
        }

        It 'Initializes module and creates vault if needed' {
            Initialize-MimecastApiModule
            Should -Invoke Register-SecretVault -Times 1
            Should -Invoke Set-ApiSecret -Times 2
        }

        It 'Skips initialization if secrets exist' {
            Mock Get-ApiSecret { ConvertTo-SecureString "existing-secret" -AsPlainText -Force }
            Initialize-MimecastApiModule
            Should -Invoke Set-ApiSecret -Times 0
        }

        It 'Forces reinitialization with -ForceSetup' {
            Mock Get-ApiSecret { ConvertTo-SecureString "existing-secret" -AsPlainText -Force }
            Initialize-MimecastApiModule -ForceSetup
            Should -Invoke Set-ApiSecret -Times 2
        }
    }

    Context 'API Connection' {
        BeforeAll {
            # Mock Get-ApiSecret to return SecureString
            Mock Get-ApiSecret { 
                ConvertTo-SecureString "test-secret" -AsPlainText -Force 
            }
            
            # Mock Invoke-RestMethod for OAuth2 token
            Mock Invoke-RestMethod {
                @{
                    access_token = 'test-token'
                    expires_in = 3600
                    token_type = 'Bearer'
                }
            } -ParameterFilter { $Uri -like '*/oauth/token' }
            
            # Mock Invoke-ApiRequest for system info
            Mock Invoke-ApiRequest {
                @{
                    success = $true
                    data = @{
                        region = 'US'
                        version = '2.0'
                        up = $true
                    }
                }
            }
        }

        It 'Connects successfully with valid OAuth2 credentials' {
            { Connect-MimecastApi -Region 'US' } | Should -Not -Throw
            
            $config = Get-MimecastApiConfiguration
            $config.Connection.Connected | Should -Be $true
            $config.Connection.Region | Should -Be 'US'
            $config.Connection.TokenExpiresAt | Should -Not -BeNullOrEmpty
        }

        It 'Connects successfully with Setup parameter' {
            Mock Get-ApiSecret { $null } # Simulate no existing credentials
            { Connect-MimecastApi -Setup -Region 'US' } | Should -Not -Throw
            Should -Invoke Set-ApiSecret -Times 2
            
            $config = Get-MimecastApiConfiguration
            $config.Connection.Connected | Should -Be $true
            $config.Connection.Region | Should -Be 'US'
        }

        It 'Uses provided credential names over defaults' {
            { Connect-MimecastApi -ClientIdName 'CustomId' -ClientSecretName 'CustomSecret' -Region 'US' } |
                Should -Not -Throw
            Should -Invoke Get-ApiSecret -ParameterFilter { $Name -eq 'CustomId' } -Times 1
            Should -Invoke Get-ApiSecret -ParameterFilter { $Name -eq 'CustomSecret' } -Times 1
        }

        It 'Gets system info with Bearer token' {
            $info = Get-MimecastApiSystemInfo
            $info.success | Should -Be $true
            $info.data.up | Should -Be $true
            
            # Verify Bearer token is used
            Should -Invoke Invoke-ApiRequest -ParameterFilter {
                $Headers['Authorization'] -like 'Bearer *'
            }
        }

        It 'Disconnects and clears OAuth2 tokens' {
            Disconnect-MimecastApi
            $config = Get-MimecastApiConfiguration
            $config.Connection | Should -Be $null
            $config.AccessToken | Should -BeNullOrEmpty
            $config.TokenExpiresAt | Should -Be $null
        }
    }
}