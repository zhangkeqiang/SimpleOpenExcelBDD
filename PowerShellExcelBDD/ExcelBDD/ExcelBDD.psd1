#
# Module manifest for module 'ExcelBDD'
#
# Generated by: Mike Zhang
#
# Generated on: 2021/6/11
#

@{

    # Script module or binary module file associated with this manifest.
    RootModule        = 'ExcelBDD.psm1'

    # Version number of this module.
    ModuleVersion     = '0.5.0'

    # ID used to uniquely identify this module
    GUID              = 'bdf35f8c-197a-4256-86ff-35da93432ba0'

    # Author of this module
    Author            = 'ExcelBDD Team'

    # Company or vendor of this module
    CompanyName       = 'ExcelBDD'

    # Copyright statement for this module
    Copyright         = 'Copyright (c) 2021 by ExcelBDD Team, licensed under Apache 2.0 License.'

    # Description of the functionality provided by this module
    Description       = 'Use Excel file as BDD feature file, get example list and testcase list from excel file'

    # Minimum version of the Windows PowerShell engine required by this module
    # PowerShellVersion = '3.0'

    TypesToProcess    = @()

    # Functions to export from this module
    FunctionsToExport = "Get-ExampleList","Get-ExampleListByHeader","Get-DataTable","Show-ExampleList"

    # # Cmdlets to export from this module
    CmdletsToExport   = @()

    # Variables to export from this module
    VariablesToExport = @()

    # # Aliases to export from this module
    AliasesToExport   = @()

    RequiredModules = "ImportExcel", "Pester"

    # List of all modules packaged with this module
    # ModuleList = @()

    # List of all files packaged with this module
    # FileList = @()

    PrivateData       = @{
        #Category of this module
        Category     = 'Scripting Techniques'

        #IsPrerelease of this module
        IsPrerelease = 'False'

        # PSData is module packaging and gallery metadata embedded in PrivateData
        # It's for rebuilding PowerShellGet (and PoshCode) NuGet-style packages
        # We had to do this because it's the only place we're allowed to extend the manifest
        # https://connect.microsoft.com/PowerShell/feedback/details/421837
        PSData       = @{
            # Keyword tags to help users find this module via navigations and search.
            Tags         = @('powershell', 'unit_testing', 'bdd', 'tdd', 'Excel')

            # The web address of an icon which can be used in galleries to represent this module
            # IconUri      = ''

            # The web address of this module's project or support homepage.
            ProjectUri   = "https://dev.azure.com/simplopen/ExcelBDD/_wiki/wikis/ExcelBDD.wiki/39/ExcelBDD-Homepage"

            # The web address of this module's license. Points to a page that's embeddable and linkable.
            LicenseUri   = "https://www.apache.org/licenses/LICENSE-2.0.html"

            # Release notes for this particular version of the module
            ReleaseNotes = 'read data from duplicated headers'
        }
    }
}
