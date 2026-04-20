param(
    [string]$TemplatePath = "c:\Project\Powerpoint\HFP\FDA_PP_Final_Use_This - 16x9 version.pptx",
    [string]$OutputPath = "c:\Project\Powerpoint\HFP\UnifiedHPC-Docs\FDA_Unified_HPC_Structured_Deck.pptx"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Escape-Xml {
    param([string]$Text)
    if ($null -eq $Text) { return "" }
    return [System.Security.SecurityElement]::Escape($Text)
}

function New-ParagraphXml {
    param(
        [string]$Text,
        [int]$FontSize = 1800,
        [bool]$Bold = $false,
        [bool]$Italic = $false,
        [string]$Color = "212121"
    )

    $escaped = Escape-Xml $Text
    $boldVal = if ($Bold) { "1" } else { "0" }
    $italicVal = if ($Italic) { "1" } else { "0" }

    return @"
<a:p>
  <a:pPr><a:buNone/></a:pPr>
  <a:r>
    <a:rPr lang="en-US" sz="$FontSize" b="$boldVal" i="$italicVal">
      <a:solidFill><a:srgbClr val="$Color"/></a:solidFill>
      <a:latin typeface="Arial"/>
    </a:rPr>
    <a:t>$escaped</a:t>
  </a:r>
  <a:endParaRPr lang="en-US" sz="$FontSize">
    <a:solidFill><a:srgbClr val="$Color"/></a:solidFill>
    <a:latin typeface="Arial"/>
  </a:endParaRPr>
</a:p>
"@
}

function New-BlankParagraphXml {
    return @"
<a:p>
  <a:pPr><a:buNone/></a:pPr>
  <a:endParaRPr lang="en-US" sz="1200">
    <a:solidFill><a:srgbClr val="212121"/></a:solidFill>
    <a:latin typeface="Arial"/>
  </a:endParaRPr>
</a:p>
"@
}

function New-SlideXml {
    param(
        [int]$SlideNumber,
        [hashtable]$Slide
    )

    $paragraphs = New-Object System.Collections.Generic.List[string]
    $paragraphs.Add((New-ParagraphXml -Text ("Title: " + $Slide.Title) -FontSize 2400 -Bold $true -Color "1F4E79"))
    $paragraphs.Add((New-ParagraphXml -Text ("Subtitle: " + $Slide.Subtitle) -FontSize 1600 -Italic $true -Color "4F4F4F"))
    $paragraphs.Add((New-BlankParagraphXml))
    $paragraphs.Add((New-ParagraphXml -Text "Bullets:" -FontSize 1800 -Bold $true -Color "1F4E79"))

    foreach ($bullet in $Slide.Bullets) {
        $paragraphs.Add((New-ParagraphXml -Text ("- " + $bullet) -FontSize 1600))
    }

    $paragraphs.Add((New-BlankParagraphXml))
    $paragraphs.Add((New-ParagraphXml -Text "Speaker Notes:" -FontSize 1800 -Bold $true -Color "1F4E79"))

    foreach ($noteLine in ($Slide.Notes -split "\r?\n")) {
        if ([string]::IsNullOrWhiteSpace($noteLine)) {
            $paragraphs.Add((New-BlankParagraphXml))
        } else {
            $paragraphs.Add((New-ParagraphXml -Text $noteLine.Trim() -FontSize 1500 -Color "333333"))
        }
    }

    $bodyXml = ($paragraphs -join "")

    return @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld name="Slide $SlideNumber">
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="0" cy="0"/>
          <a:chOff x="0" y="0"/>
          <a:chExt cx="0" cy="0"/>
        </a:xfrm>
      </p:grpSpPr>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="2" name="ContentBox"/>
          <p:cNvSpPr txBox="1"/>
          <p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="457200" y="304800"/>
            <a:ext cx="8229600" cy="4419600"/>
          </a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          <a:noFill/>
        </p:spPr>
        <p:txBody>
          <a:bodyPr wrap="square">
            <a:spAutoFit/>
          </a:bodyPr>
          <a:lstStyle/>
          $bodyXml
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>
"@
}

function New-PictureSlideXml {
    param(
        [int]$SlideNumber
    )

    return @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld name="Slide $SlideNumber">
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="0" cy="0"/>
          <a:chOff x="0" y="0"/>
          <a:chExt cx="0" cy="0"/>
        </a:xfrm>
      </p:grpSpPr>
      <p:pic>
        <p:nvPicPr>
          <p:cNvPr id="2" name="SlideImage"/>
          <p:cNvPicPr/>
          <p:nvPr/>
        </p:nvPicPr>
        <p:blipFill>
          <a:blip r:embed="rId2"/>
          <a:stretch>
            <a:fillRect/>
          </a:stretch>
        </p:blipFill>
        <p:spPr>
          <a:xfrm>
            <a:off x="0" y="0"/>
            <a:ext cx="9144000" cy="5143500"/>
          </a:xfrm>
          <a:prstGeom prst="rect">
            <a:avLst/>
          </a:prstGeom>
        </p:spPr>
      </p:pic>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>
"@
}

function New-SlideRelsXml {
    param(
        [bool]$HasImage,
        [string]$ImageTarget
    )

    $rels = @(
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
        '  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>'
    )

    if ($HasImage) {
        $rels += ('  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/' + $ImageTarget + '"/>')
    }

    $rels += '</Relationships>'
    return ($rels -join "`r`n")
}

function Write-ZipWithForwardSlashes {
    param(
        [string]$SourceDirectory,
        [string]$DestinationZip
    )

    if (Test-Path -LiteralPath $DestinationZip) {
        Remove-Item -LiteralPath $DestinationZip -Force
    }

    $fileStream = [System.IO.File]::Open($DestinationZip, [System.IO.FileMode]::CreateNew)
    try {
        $zip = New-Object System.IO.Compression.ZipArchive($fileStream, [System.IO.Compression.ZipArchiveMode]::Create)
        try {
            $basePath = (Resolve-Path $SourceDirectory).Path
            $baseLength = $basePath.Length + 1
            Get-ChildItem -LiteralPath $SourceDirectory -Recurse -File | ForEach-Object {
                $relativePath = $_.FullName.Substring($baseLength).Replace("\", "/")
                [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, $_.FullName, $relativePath, [System.IO.Compression.CompressionLevel]::Optimal) | Out-Null
            }
        } finally {
            $zip.Dispose()
        }
    } finally {
        $fileStream.Dispose()
    }
}

$textSlides = @(
    @{
        Title = "FDA Unified HPC: A Service-Based, Agency-Wide Model Built on Existing Infrastructure"
        Subtitle = "Create a unified HPC capability that all FDA centers can use"
        Bullets = @(
            "Agency-wide shared HPC service",
            "Built on existing center infrastructure",
            "Centers retain ownership of local clusters",
            "Improves access, collaboration, and resource sharing"
        )
        Notes = "FDA already has meaningful HPC investments across multiple centers, including CDRH, HIVE/NCTR, and HFP environments. The opportunity is to connect and operationalize those assets as an agency-wide service so scientists can use them more easily and effectively. This is a proposal to unify access and strengthen the value of infrastructure FDA already owns."
    },
    @{
        Title = "FDA Already Has Significant HPC Assets"
        Subtitle = "Existing on-premise capacity, GPU resources, and center expertise already exist"
        Bullets = @(
            "Significant on-premise HPC capacity",
            "CPU, GPU, storage, and workflow assets exist today",
            "Strong center-specific capabilities already in place",
            "Main challenge is fragmentation"
        )
        Notes = "FDA is not starting from zero. Multiple centers have already developed strong HPC environments that support genomics, analytics, AI/ML, and scientific computing. The issue is that those assets are distributed and managed independently, which limits their agency-wide value. The data call materials point to substantial shared infrastructure already in place, including hundreds of compute nodes, thousands of cores, and GPU resources across FDA environments."
    },
    @{
        Title = "Current State: Powerful but Distributed and Uneven"
        Subtitle = "Different environments, tools, and access patterns limit collaboration"
        Bullets = @(
            "Powerful but distributed environments",
            "Different access methods and tools",
            "Data often tied to local clusters",
            "Uneven user experience and access across centers"
        )
        Notes = "Each center has built what it needed for its mission, but the result is a patchwork environment. Authentication, workflow tooling, storage patterns, and support models differ from place to place. That makes it harder for scientists to move quickly and harder for FDA to manage capacity as a shared enterprise asset. Some centers already have more user-friendly web entry points, while others still lean heavily on command-line access."
    },
    @{
        Title = "Why Siloed HPC Creates Agency-Wide Friction"
        Subtitle = "Fragmentation makes access, sharing, and support harder than they should be"
        Bullets = @(
            "Hard to discover and access available resources",
            "Difficult to share data and workloads",
            "Duplicate administration and support",
            "Idle capacity in one area cannot easily help another"
        )
        Notes = "The siloed model creates both user pain and operational inefficiency. Scientists must navigate local rules and systems, while administrators across centers solve similar problems separately. At the same time, unused capacity in one environment is not easily available to another group that needs it. That is why the agency is trying to move from center-by-center ownership to a shared-service experience."
    },
    @{
        Title = "Key Problems the Unified HPC Initiative Addresses"
        Subtitle = "FDA needs simpler access, shared storage, and better visibility"
        Bullets = @(
            "No single FDA HPC entry point",
            "Limited shared storage and orchestration",
            "Inconsistent user experience",
            "Rising cloud GPU costs complicate cloud-only strategies",
            "Limited visibility into utilization and cost"
        )
        Notes = "The initiative is focused on solving a defined set of enterprise problems: access, storage, usability, visibility, coordination, and affordability. These are practical issues that affect delivery speed, support burden, and the ability to scale emerging workloads such as AI and advanced analytics."
    },
    @{
        Title = "Unified HPC Is a Shared-Service Model, Not a Forced Consolidation"
        Subtitle = "Centers retain local systems while FDA adds a common service layer"
        Bullets = @(
            "Shared service, not forced consolidation",
            "Centers retain local systems and expertise",
            "FDA adds a common service layer",
            "Scientists gain broader resource access"
        )
        Notes = "Unified HPC does not mean taking away center-owned systems. It means introducing a common operating model that makes those systems work better together. The emphasis is on interoperability, agency-wide discoverability, and shared services while preserving local mission ownership. The goal is a federated experience with a unified front door, not a single monolithic cluster."
    },
    @{
        Title = "How We Unify HPC Using Current Infrastructure"
        Subtitle = "Start with shared storage, then add common access across existing clusters"
        Bullets = @(
            "Start with shared storage as the first priority",
            "Connect existing on-premise clusters",
            "Add common portal and workflows",
            "Support current schedulers and tools"
        )
        Notes = "The implementation path should begin with the infrastructure pieces that unlock the most value early. Shared storage is the first critical step because it enables collaboration and workload mobility, and it was identified in the discussion as the top priority. A common portal, workflow patterns, and identity-aware access can then make the environment usable at scale without forcing immediate redesign of every local system."
    },
    @{
        Title = "What a Unified FDA HPC Service Should Provide"
        Subtitle = "Single access, shared storage, workflows, accounting, and guided interfaces"
        Bullets = @(
            "Single user entry point for all FDA HPC services",
            "Shared storage and policy-based data movement across centers",
            "Standard workflow templates for genomics, AI/ML, modeling, and analytics",
            "Common accounting for compute, GPU, storage, and cloud usage",
            "Better support for non-technical users through graphical interfaces and guided workflows",
            "Secure support for private AI/LLM workloads and future Kubernetes-based services where appropriate"
        )
        Notes = "This is a pragmatic strategy. It improves return on existing investments first, instead of leading with a large rebuild or broad cloud migration. Cloud can still play a role, but in a targeted way for surge demand or specialized workloads where it provides clear value. The point is to make the service easy to use regardless of the user's technical skill level."
    },
    @{
        Title = "Benefits of a Unified Service-Based HPC Model"
        Subtitle = "Faster science, better utilization, and stronger collaboration"
        Bullets = @(
            "Faster time to science",
            "Better utilization of existing FDA infrastructure before buying more hardware",
            "Lower duplication of operational effort across centers",
            "Stronger cost control through shared visibility and selective cloud use",
            "Better platform for AI, analytics, and future regulatory science workloads"
        )
        Notes = "The benefits are both operational and strategic. Users get simpler access and faster movement from data to results. FDA gets better visibility, more flexible use of existing assets, and a stronger base for future computational needs, especially in AI, advanced analytics, and shared mission services. Better shared storage and clearer service boundaries also support collaboration across centers."
    },
    @{
        Title = "Recommendation and Next Steps"
        Subtitle = "Launch a focused pilot and scale based on results"
        Bullets = @(
            "Endorse a unified, service-based FDA HPC model built on existing center infrastructure",
            "Begin with an implementation pilot focused on shared storage, common access, and 2-3 representative clusters",
            "Develop agency-wide governance, a service catalog, and success metrics",
            "Collect center-specific requirements for hardware, storage, AI, and networking priorities",
            "Use the pilot to validate cost, adoption, security, and operational impact before full expansion"
        )
        Notes = "The best next step is a focused pilot that proves the concept in FDA's real environment. That pilot should validate usability, security, cost visibility, and workload portability. If successful, it provides the evidence and operating model needed for phased expansion across the agency. This is the moment to move from discussion to a shared implementation path."
    }
)

$imageSlides = @{
    1 = "slide1_unified_hpc.png"
    4 = "slide4_siloed_hpc.png"
    8 = "slide8_service_value.png"
}

$deckPlan = New-Object System.Collections.Generic.List[hashtable]
for ($i = 0; $i -lt $textSlides.Count; $i++) {
    $slideNumber = $i + 1
    $deckPlan.Add(@{
        Kind = "text"
        SourceSlideNumber = $slideNumber
        Slide = $textSlides[$i]
    })

    if ($imageSlides.ContainsKey($slideNumber)) {
        $deckPlan.Add(@{
            Kind = "image"
            SourceSlideNumber = $slideNumber
            ImageFile = $imageSlides[$slideNumber]
        })
    }
}

$slideCount = $deckPlan.Count

if (-not (Test-Path -LiteralPath $TemplatePath)) {
    throw "Template file not found: $TemplatePath"
}

$tempRoot = Join-Path $env:TEMP ("UnifiedHPCPpt_" + [guid]::NewGuid().ToString("N"))
$extractDir = Join-Path $tempRoot "extract"
$zipPath = Join-Path $tempRoot "deck.zip"
$mediaDir = Join-Path $extractDir "ppt/media"
$docsRoot = Split-Path -Parent $OutputPath

New-Item -ItemType Directory -Path $extractDir -Force | Out-Null
New-Item -ItemType Directory -Path $mediaDir -Force | Out-Null

Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.IO.Compression.ZipFile]::ExtractToDirectory($TemplatePath, $extractDir)

for ($i = 1; $i -le $slideCount; $i++) {
    $slidePath = Join-Path $extractDir ("ppt/slides/slide{0}.xml" -f $i)
    $slideRelsPath = Join-Path $extractDir ("ppt/slides/_rels/slide{0}.xml.rels" -f $i)
    $planEntry = $deckPlan[$i - 1]

    if ($planEntry.Kind -eq "image") {
        $slideXml = New-PictureSlideXml -SlideNumber $i
        $slideRels = New-SlideRelsXml -HasImage $true -ImageTarget $planEntry.ImageFile
        $sourceImagePath = Join-Path $docsRoot "Slide_Images\$($planEntry.ImageFile)"
        if (-not (Test-Path -LiteralPath $sourceImagePath)) {
            throw "Image slide asset not found: $sourceImagePath"
        }
        Copy-Item -LiteralPath $sourceImagePath -Destination (Join-Path $mediaDir $planEntry.ImageFile) -Force
    } else {
        $slideXml = New-SlideXml -SlideNumber $i -Slide $planEntry.Slide
        $slideRels = New-SlideRelsXml -HasImage $false -ImageTarget $null
    }

    [System.IO.File]::WriteAllText($slidePath, $slideXml, [System.Text.UTF8Encoding]::new($false))
    New-Item -ItemType Directory -Path (Split-Path -Parent $slideRelsPath) -Force | Out-Null
    [System.IO.File]::WriteAllText($slideRelsPath, $slideRels, [System.Text.UTF8Encoding]::new($false))
}

$slideDir = Join-Path $extractDir "ppt/slides"
Get-ChildItem -LiteralPath $slideDir -Filter "slide*.xml" | Where-Object {
    if ($_.Name -match '^slide(\d+)\.xml$') {
        [int]$Matches[1] -gt $slideCount
    } else {
        $false
    }
} | Remove-Item -Force

$slideRelDir = Join-Path $slideDir "_rels"
if (Test-Path -LiteralPath $slideRelDir) {
    Get-ChildItem -LiteralPath $slideRelDir -Filter "slide*.xml.rels" | Where-Object {
        if ($_.Name -match '^slide(\d+)\.xml\.rels$') {
            [int]$Matches[1] -gt $slideCount
        } else {
            $false
        }
    } | Remove-Item -Force
}

$presentationPath = Join-Path $extractDir "ppt/presentation.xml"
[xml]$presentationXml = Get-Content -LiteralPath $presentationPath
$ns = New-Object System.Xml.XmlNamespaceManager($presentationXml.NameTable)
$ns.AddNamespace("p", "http://schemas.openxmlformats.org/presentationml/2006/main")
$sldIdLst = $presentationXml.SelectSingleNode("//p:sldIdLst", $ns)
while ($sldIdLst.ChildNodes.Count -gt $slideCount) {
    $sldIdLst.RemoveChild($sldIdLst.LastChild) | Out-Null
}
while ($sldIdLst.ChildNodes.Count -lt $slideCount) {
    $newId = 441 + $sldIdLst.ChildNodes.Count
    $newSlideId = $presentationXml.CreateElement("p", "sldId", "http://schemas.openxmlformats.org/presentationml/2006/main")
    $newSlideId.SetAttribute("id", [string]$newId)
    $sldIdLst.AppendChild($newSlideId) | Out-Null
}
$sldIdNodes = @($sldIdLst.ChildNodes)
for ($i = 0; $i -lt $slideCount; $i++) {
    $sldIdNode = $sldIdNodes[$i]
    $null = $sldIdNode.SetAttribute("id", [string](441 + $i))
    $slideRelId = switch ($i) {
        0  { "rId5" }
        1  { "rId6" }
        2  { "rId7" }
        3  { "rId8" }
        4  { "rId9" }
        5  { "rId10" }
        6  { "rId11" }
        7  { "rId12" }
        8  { "rId13" }
        9  { "rId14" }
        10 { "rId15" }
        11 { "rId16" }
        12 { "rId24" }
        default { throw "Unexpected slide index: $i" }
    }
    $null = $sldIdNode.SetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships", $slideRelId)
}
$presentationXml.Save($presentationPath)

$presentationRelsPath = Join-Path $extractDir "ppt/_rels/presentation.xml.rels"
[xml]$presentationRelsXml = Get-Content -LiteralPath $presentationRelsPath
$relsNs = New-Object System.Xml.XmlNamespaceManager($presentationRelsXml.NameTable)
$relsNs.AddNamespace("pr", "http://schemas.openxmlformats.org/package/2006/relationships")
$slideRelNodes = @($presentationRelsXml.SelectNodes("//pr:Relationship[contains(@Type,'/slide')]", $relsNs))
foreach ($node in $slideRelNodes) {
    $node.ParentNode.RemoveChild($node) | Out-Null
}

$slideRelTargets = @(
    "slides/slide1.xml",
    "slides/slide2.xml",
    "slides/slide3.xml",
    "slides/slide4.xml",
    "slides/slide5.xml",
    "slides/slide6.xml",
    "slides/slide7.xml",
    "slides/slide8.xml",
    "slides/slide9.xml",
    "slides/slide10.xml",
    "slides/slide11.xml",
    "slides/slide12.xml",
    "slides/slide13.xml"
)

for ($i = 0; $i -lt $slideRelTargets.Count; $i++) {
    $rel = $presentationRelsXml.CreateElement("Relationship", $presentationRelsXml.DocumentElement.NamespaceURI)
    $relId = switch ($i) {
        0  { "rId5" }
        1  { "rId6" }
        2  { "rId7" }
        3  { "rId8" }
        4  { "rId9" }
        5  { "rId10" }
        6  { "rId11" }
        7  { "rId12" }
        8  { "rId13" }
        9  { "rId14" }
        10 { "rId15" }
        11 { "rId16" }
        12 { "rId24" }
        default { throw "Unexpected slide relationship index: $i" }
    }
    $rel.SetAttribute("Id", $relId)
    $rel.SetAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide")
    $rel.SetAttribute("Target", $slideRelTargets[$i])
    $presentationRelsXml.DocumentElement.AppendChild($rel) | Out-Null
}
$presentationRelsXml.Save($presentationRelsPath)

$contentTypesPath = Join-Path $extractDir "[Content_Types].xml"
[xml]$contentTypesXml = Get-Content -LiteralPath $contentTypesPath
$contentNs = New-Object System.Xml.XmlNamespaceManager($contentTypesXml.NameTable)
$contentNs.AddNamespace("ct", "http://schemas.openxmlformats.org/package/2006/content-types")
$slideOverrides = @($contentTypesXml.SelectNodes("//ct:Override[starts-with(@PartName, '/ppt/slides/slide')]", $contentNs))
foreach ($node in $slideOverrides) {
    $node.ParentNode.RemoveChild($node) | Out-Null
}
for ($i = 1; $i -le $slideCount; $i++) {
    $override = $contentTypesXml.CreateElement("Override", $contentTypesXml.DocumentElement.NamespaceURI)
    $override.SetAttribute("PartName", "/ppt/slides/slide{0}.xml" -f $i)
    $override.SetAttribute("ContentType", "application/vnd.openxmlformats-officedocument.presentationml.slide+xml")
    $contentTypesXml.DocumentElement.AppendChild($override) | Out-Null
}
$contentTypesXml.Save($contentTypesPath)

$appPropsPath = Join-Path $extractDir "docProps/app.xml"
if (Test-Path -LiteralPath $appPropsPath) {
    [xml]$appXml = Get-Content -LiteralPath $appPropsPath
    $propsNs = New-Object System.Xml.XmlNamespaceManager($appXml.NameTable)
    $propsNs.AddNamespace("ep", "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties")
    $propsNs.AddNamespace("vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")
    $slidesNode = $appXml.SelectSingleNode("//*[local-name()='Slides']")
    if ($slidesNode) {
        $slidesNode.InnerText = [string]$slideCount
    }
    $headingPairsNode = $appXml.SelectSingleNode("//*[local-name()='HeadingPairs']")
    if ($headingPairsNode) {
        $headingCountNode = $headingPairsNode.SelectNodes(".//*[local-name()='i4']")
        if ($headingCountNode -and $headingCountNode.Count -ge 2) {
            $headingCountNode.Item(1).InnerText = [string]$slideCount
        }
    }
    $appXml.Save($appPropsPath)
}

if (Test-Path -LiteralPath $OutputPath) {
    Remove-Item -LiteralPath $OutputPath -Force
}

Write-ZipWithForwardSlashes -SourceDirectory $extractDir -DestinationZip $zipPath
Move-Item -LiteralPath $zipPath -Destination $OutputPath

Remove-Item -LiteralPath $tempRoot -Recurse -Force

Write-Output "Created: $OutputPath"
