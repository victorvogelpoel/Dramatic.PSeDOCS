


Import-Module Pester


Describe 'Describe line' {

    BeforeAll {
        Write-Host 'BeforeAll at Describe scope'
    }

    AfterAll {
        Write-Host 'AfterAll at Describe scope'
    }

    BeforeEach {
        Write-Host 'BeforeEach at Describe scope'
    }

    AfterEach {
        Write-Host 'AfterEach at Describe scope'
    }

    Context 'Context line' {

        BeforeAll {
            Write-Host 'BeforeAll at Context scope'
        }

        AfterAll {
            Write-Host 'AfterAll at Context scope'
        }

        BeforeEach {
            Write-Host 'BeforeEach at Context scope'
        }

        AfterEach {
            Write-Host 'AfterEach at Context scope'
        }

        It 'It line' {
            $true | Should Be $true
        }

        It 'Second it line' {
            $true | Should Be $true
        }
    }

    It 'Third it line' {
        $true | Should Be $true
    }
}

<# RESULT
    Describing Describe line
    BeforeAll at Describe scope

      Context Context line
    BeforeAll at Context scope
    BeforeEach at Describe scope
    BeforeEach at Context scope
    AfterEach at Context scope
    AfterEach at Describe scope
        [+] It line 74ms

    BeforeEach at Describe scope
    BeforeEach at Context scope
    AfterEach at Context scope
    AfterEach at Describe scope
        [+] Second it line 19ms

    AfterAll at Context scope
    BeforeEach at Describe scope
    AfterEach at Describe scope
      [+] Third it line 37ms

    AfterAll at Describe scope
#>
