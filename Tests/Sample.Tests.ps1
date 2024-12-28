Describe 'Sample Pester Tests' {
    It 'should return true for a true condition' {
        $true | Should -Be $true
    }
    
    It 'should return false for a false condition' {
        $false | Should -Be $false
    }
}