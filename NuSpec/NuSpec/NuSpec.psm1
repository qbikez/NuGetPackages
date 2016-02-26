function Get-SolutionDir {
    if($dte.Solution -and $dte.Solution.IsOpen) {
        return Split-Path $dte.Solution.Properties.Item("Path").Value
    }
    else {
        throw "Solution not avaliable"
    }
}

function Resolve-ProjectName {
    param(
        [parameter(ValueFromPipelineByPropertyName = $true)]
        [string[]]$ProjectName,
        [Switch]
        [bool]$All = $true
    )
    
    if($ProjectName) {
        $projects = Get-Project $ProjectName
    }
    else {
        # All projects by default
        $projects = Get-Project -All:$All
    }
    
    $projects
}

function Get-MSBuildProject {
    param(
        [parameter(ValueFromPipelineByPropertyName = $true)]
        [string[]]$ProjectName
    )
    Process {
        (Resolve-ProjectName $ProjectName) | % {
            $path = $_.FullName
            @([Microsoft.Build.Evaluation.ProjectCollection]::GlobalProjectCollection.GetLoadedProjects($path))[0]
        }
    }
}

function Set-MSBuildProperty {
    param(
        [parameter(Position = 0, Mandatory = $true)]
        $PropertyName,
        [parameter(Position = 1, Mandatory = $true)]
        $PropertyValue,
        [parameter(Position = 2, ValueFromPipelineByPropertyName = $true)]
        [string[]]$ProjectName
    )
    Process {
        (Resolve-ProjectName $ProjectName) | %{
            $buildProject = $_ | Get-MSBuildProject
            $buildProject.SetProperty($PropertyName, $PropertyValue) | Out-Null
            $_.Save()
        }
    }
}

function Get-MSBuildProperty {
    param(
        [parameter(Position = 0, Mandatory = $true)]
        $PropertyName,
        [parameter(Position = 2, ValueFromPipelineByPropertyName = $true)]
        [string]$ProjectName
    )
    
    $buildProject = Get-MSBuildProject $ProjectName
    $buildProject.GetProperty($PropertyName)
}

function Install-NuSpec {
    param(
        [parameter(ValueFromPipelineByPropertyName = $true)]
        [string[]]$ProjectName,
    	[switch]$EnableIntelliSense,
        [string]$TemplatePath
    )
    
    Process {
    
        $projects = (Resolve-ProjectName $ProjectName)
        
        if(!$projects) {
            Write-Error "Unable to locate project. Make sure it isn't unloaded."
            return
        }
		
        $m = get-module NuSpec
		#$profileDirectory = Split-Path $profile -parent
		#$profileModulesDirectory = (Join-Path $profileDirectory "Modules")
		$moduleDir = split-path -parent $m.Path
		
        if($EnableIntelliSense){
            Enable-NuSpecIntelliSense            
        }
        
        # Add NuSpec file for project(s)
        $projects | %{ 
            $project = $_
            
            # Set the nuspec target path
            $projectFile = Get-Item $project.FullName
            $projectDir = [System.IO.Path]::GetDirectoryName($projectFile)
            $projectNuspec = "$($project.Name).nuspec"
            $projectNuspecPath = Join-Path $projectDir $projectNuspec
            
            # Get the nuspec template source path
            if($TemplatePath) {
                $nuspecTemplatePath = $TemplatePath
            }
            else {
                $nuspecTemplatePath = Join-Path $moduleDir NuSpecTemplate.xml
            }
            
            write-verbose "creating nuspec at '$projectNuspecPath' from template '$nuspecTemplatePath'"
            # Copy the templated nuspec to the project nuspec if it doesn't exist
            if(!(Test-Path $projectNuspecPath)) {
                Copy-Item $nuspecTemplatePath $projectNuspecPath
            }
            else {
                Write-Warning "Failed to install nuspec '$projectNuspec' into '$($project.Name)' because the file already exists."
            }
            
            try {
                write-verbose "adding nuspec to project"
                # Add nuspec file to the project
                add-itemToProject $project $projectNuspecPath
				
				Set-MSBuildProperty NuSpecFile $projectNuspec $project.Name
                
                "Updated '$($project.Name)' to use nuspec '$projectNuspec'"
            }
            catch {
                Write-Warning "Failed to install nuspec '$projectNuspec' into '$($project.Name)': $_"
            }
        }
    }
}

function Enable-NuSpecIntelliSense {
    Process {		
		$profileDirectory = Split-Path $profile -parent
		$profileModulesDirectory = (Join-Path $profileDirectory "Modules")
		$moduleDir = (Join-Path $profileModulesDirectory "NuSpec")

        $solutionDir = Get-SolutionDir
        $solution = Get-Interface $dte.Solution ([EnvDTE80.Solution2])
        
        # Set up solution folder "Solution Items"
        $solutionItemsProject = $dte.Solution.Projects | Where-Object { $_.ProjectName -eq "Solution Items" }
        if(!($solutionItemsProject)) {
            $solutionItemsProject = $solution.AddSolutionFolder("Solution Items")
        }        
        
        # Copy the XSD in the solution directory
        try {
            $xsdInstallPath = Join-Path $solutionDir 'nuspec.xsd'
            $xsdToolsPath = Join-Path $moduleDir 'nuspec.xsd'
                
            if(!(Test-Path $xsdInstallPath)) {
                Copy-Item $xsdToolsPath $xsdInstallPath
            }
                
            $alreadyAdded = $solutionItemsProject.ProjectItems | Where-Object { $_.Name -eq 'nuspec.xsd' }
            if(!($alreadyAdded)) {
                $solutionItemsProject.ProjectItems.AddFromFile($xsdInstallPath) | Out-Null
            }
        }
        catch {
            Write-Warning "Failed to install nuspec.xsd into 'Solution Items'"
        }
        $solution.SaveAs($solution.FullName)
    }
}

function Pack-Nuget {
    param (
        [parameter(ValueFromPipelineByPropertyName = $true)]
        [string[]]$ProjectName,
        [string] $specFile,
        [switch][bool] $noProjectRefs
    )
    Process {

        $projects = (Resolve-ProjectName $ProjectName)
        
        if(!$projects) {
            Write-Error "Unable to locate project. Make sure it isn't unloaded."
            return
        }
		
		$profileDirectory = Split-Path $profile -parent
		$profileModulesDirectory = (Join-Path $profileDirectory "Modules")
		$moduleDir = (Join-Path $profileModulesDirectory "NuSpec")
      
        
        $projects | % { 
            Push-Location
            try {
                $project = $_
            
                # Set the nuspec target path
                $projectFile = Get-Item $project.FullName
                $projectDir = [System.IO.Path]::GetDirectoryName($projectFile)
                $projectNuspec = "$($project.Name).nuspec"
                if (![string]::isnullorempty($specFile)) {
					$projectNuspec = $specFile
                }
                $projectNuspecPath = Join-Path $projectDir $projectNuspec
                
                cd $projectDir
                $args = @()


                if (![string]::isnullorempty($specFile)) {
                    $args += $projectNuspecPath
                } else {
                    $args += $projectFile
                }

                if (!$noProjectRefs) {
                    $args += "-IncludeReferencedProjects"
                }
                
                Write-Warning "running command 'nuget pack $args'"
                nuget pack @args

                
            }
            finally {
                Pop-Location
            }
        }      
    }
}

function Push-Nuget {
    param (
        [parameter(ValueFromPipelineByPropertyName = $true)]
        [string[]]$ProjectName,
        [string] $specFile,
        [parameter(Mandatory=$true)]
        [string] $Source,
        [switch][bool] $Pack = $true,
        [switch][bool] $noProjectRefs = $false
    )
    Process {

        $projects = (Resolve-ProjectName $ProjectName)
        
        if(!$projects) {
            Write-Error "Unable to locate project. Make sure it isn't unloaded."
            return
        }

        if ($Pack) {
            Pack-Nuget -ProjectName $ProjectName -specFile $specFile -noProjectRefs:$noProjectRefs
        }
		
		$profileDirectory = Split-Path $profile -parent
		$profileModulesDirectory = (Join-Path $profileDirectory "Modules")
		$moduleDir = (Join-Path $profileModulesDirectory "NuSpec")
      
        
        $projects | % { 
            Push-Location
            try {
                $project = $_
            
                # Set the nuspec target path
                $projectFile = Get-Item $project.FullName
                $projectDir = [System.IO.Path]::GetDirectoryName($projectFile)
                $projectNuspec = "$($project.Name).nuspec"
                if (![string]::isnullorempty($specFile)) {
					$projectNuspec = $specFile
                }
                $projectNuspecPath = Join-Path $projectDir $projectNuspec
                
                cd $projectDir
                
                

                if (![string]::isnullorempty($specFile)) {
                    $nupkgs = gci . -Filter "$specFile.*.nupkg" -File
                } else {
                    $nupkgs = gci . -Filter "*.nupkg" -File
                }

                $nupkgs = $nupkgs | sort -Descending -Property CreationTime
                $nupkgs 
                $pkg = $nupkgs[0]
                
                nuget push $pkg.FullName -source $source
            }
            finally {
                Pop-Location
            }
        }      
    }
}

function add-itemToProject($project, $file) {
    if ($project.ProjectItems -ne $null) {
        $project.ProjectItems.AddFromFile($file) | Out-Null
        $project.Save()
    } else {
        # assume a fallback function 'add-projectItem' exists
        add-projectItem $project.FullName $file -Verbose:$($VerbosePreference -eq "Continue")
    }
}

# Statement completion for project names
'Install-NuSpec', 'Enable-NuSpecIntelliSense', 'Pack-Nuget', 'Push-Nuget'  | %{ 
    #Register-TabExpansion $_ @{
    #    ProjectName = { Get-Project -All | Select -ExpandProperty Name }
    #}
}

Export-ModuleMember Install-NuSpec, Enable-NuSpecIntelliSense, Pack-Nuget, Push-Nuget