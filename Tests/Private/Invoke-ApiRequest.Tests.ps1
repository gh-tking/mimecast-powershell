Describe 'Invoke-ApiRequest Tests' {
    BeforeAll {
        $ModulePath = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
        . "$ModulePath\Private\Invoke-ApiRequest.ps1"
        
        # Mock module's private data
        $ExecutionContext.SessionState.Module.PrivateData = @{
            BaseUri = 'https://api.example.com'
        }
    }

    Context 'Parameter Validation' {
        It 'Requires Method parameter' {
            { Invoke-ApiRequest -Path '/test' } | 
                Should -Throw -ExpectedMessage '*Method*'
        }

        It 'Requires Path parameter' {
            { Invoke-ApiRequest -Method 'GET' } | 
                Should -Throw -ExpectedMessage '*Path*'
        }

        It 'Validates Method parameter' {
            { Invoke-ApiRequest -Method 'INVALID' -Path '/test' } | 
                Should -Throw -ExpectedMessage '*ValidateSet*'
        }
    }

    Context 'API Calls' {
        BeforeAll {
            # Mock Invoke-RestMethod
            Mock Invoke-RestMethod {
                return @{
                    status = 'success'
                    data = @{
                        id = 1
                        name = 'Test'
                    }
                }
            }
        }

        It 'Makes GET request successfully' {
            $result = Invoke-ApiRequest -Method 'GET' -Path '/test'
            $result.status | Should -Be 'success'
            Should -Invoke Invoke-RestMethod -Times 1
        }

        It 'Includes query parameters in URI' {
            $query = @{
                page = 1
                limit = 10
            }
            
            $result = Invoke-ApiRequest -Method 'GET' -Path '/test' -QueryParameters $query
            Should -Invoke Invoke-RestMethod -Times 1 -ParameterFilter {
                $Uri -like '*page=1*' -and $Uri -like '*limit=10*'
            }
        }
    }
}