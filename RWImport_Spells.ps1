Add-Type @'
public class objIDMap
{
    public string name;
    public string category_id;
    public string alias;
}
'@

Function CreateContainer($structure,$contents,$container) {
    # Create Container
    $NewTopic = $contents.OwnerDocument.CreateElement("topic",$contents.NamespaceURI)
    $NewTopic.setattribute("topic_id","Topic_$script:topic_id")
    $script:topic_id++
    $NewTopic.setattribute("public_name",$container.name)
    $NewTopic.setattribute("category_id",$container.category_id)
    $NewTopic = $contents.AppendChild($NewTopic)

    if ($container.alias -ne $null) {
        $Alias = $NewTopic.OwnerDocument.CreateElement("alias",$NewTopic.NamespaceURI)
        $Alias.setattribute("alias_id","Alias_$script:alias_id")
        $script:alias_id++
        $Alias.setattribute("name",$container.alias)
        # $Alias.setattribute("is_show_nav_pane","false")
        $Alias = $NewTopic.AppendChild($Alias)
    }
    
    if ($container.user) {
        $OverviewCategory = $structure.category | Where-Object {$_.category_id -eq $container.category_id}
        $OverviewSection = $OverviewCategory.partition | Where-Object {$_.name -eq "Overview"}
    } else {
        $OverviewCategory = $structure.category_global | Where-Object {$_.category_id -eq $container.category_id}
        $OverviewSection = $OverviewCategory.partition_global | Where-Object {$_.name -eq "Overview"}
    }

    $NewSection = $NewTopic.OwnerDocument.CreateElement("section",$NewTopic.NamespaceURI)
    $NewSection.SetAttribute("partition_id",$OverviewSection.partition_id)
    $NewSection = $NewTopic.AppendChild($NewSection)
    
    $NewSnippet = $NewSection.OwnerDocument.CreateElement("snippet",$NewSection.NamespaceURI)
    $NewSnippet.SetAttribute("type","Multi_Line")
    $NewSnippet = $NewSection.AppendChild($NewSnippet)

    Return $NewTopic
    # $NewContent = $NewSnippet.OwnerDocument.CreateElement("contents",$NewSnippet.NamespaceURI)
    # $NewContent = $NewSnippet.AppendChild($NewContent)
} # Function CreateContainer

Function ImportSpells($structure,$SpellContainer,$SpellCategory,$LetterSpells) {

    foreach ($Spell in $LetterSpells) {
        # Add the spell topic
        $NewSpell = $SpellContainer.OwnerDocument.CreateElement("topic",$SpellContainer.NamespaceURI)
        $NewSpell.setattribute("topic_id","Topic_$Script:topic_id")
        $Script:topic_id++
        $NewSpell.setattribute("public_name",$Spell.name)
        $NewSpell.setattribute("suffix",$Spell.source)
        $NewSpell.setattribute("category_id",$SpellCategory.category_id)
        $NewSpell = $SpellContainer.AppendChild($NewSpell)

        # Add all the sections to this topic
        $SectionList = $SpellCategory.GetElementsByTagName("*") | Where-Object {$_.partition_id}
        foreach ($Section in $SectionList) {
            $NewSection = $NewSpell.OwnerDocument.CreateElement("section",$NewSpell.NamespaceURI)
            $NewSection.SetAttribute("partition_id",$Section.partition_id)
            $NewSection = $NewSpell.AppendChild($NewSection)
            
            # Add all the snippet types to the current section
            Switch ($Section.name) {
                "Overview" {
                    $FacetList = $Section.GetElementsByTagName("*") | Where-Object {$_.facet_id}
                    foreach ($Facet in $FacetList) {
                        $NewSnippet = $NewSection.OwnerDocument.CreateElement("snippet",$NewSection.NameSpaceURI)
                        $NewSnippet.setattribute("facet_id",$Facet.facet_id)
                        $NewSnippet = $NewSection.AppendChild($NewSnippet)

                        Switch ($Facet.name) {
                            "School" {
                                $NewSnippet.setattribute("type","Tag_Standard")

                                # This code is in lieu of tags working properly.
                                $NewAnnotation = $NewSnippet.OwnerDocument.CreateElement("annotation",$NewSnippet.NamespaceURI)
                                $NewAnnotation = $NewSnippet.AppendChild($NewAnnotation)
                                $NewSnippet.annotation = $Spell.school


                                <# The code below is deactivated until tags are working properly.
                                $SnippetTags = $structure.domain_global | Where-Object {$_.name -eq "Spell School"}
                                $AssignTag = ($SnippetTags.GetElementsByTagName("*") | Where-Object {$_.tag_id}) | Where-Object {$_.name -eq $Spell.school}
                                if ($AssignTag) {
                                    $NewTag = $NewSnippet.OwnerDocument.CreateElement("tag_assign",$NewSnippet.NameSpaceURI)
                                    $NewTag.setattribute("tag_id",$AssignTag.tag_id)
                                    $NewTag.setattribute("type","Indirect")
                                    $NewTag = $NewSnippet.AppendChild($NewTag)
                                } elseif (($Spell.school.contains(",")) -or ($Spell.school.contains(" or ")) -or ($Spell.school.contains("see text"))) {
                                    $AssignTag = ($SnippetTags.GetElementsByTagName("*") | Where-Object {$_.tag_id}) | Where-Object {$_.name -eq "see text or annotation"}
                                    $NewAnnotation = $NewSnippet.OwnerDocument.CreateElement("annotation",$NewSnippet.NamespaceURI)
                                    $NewAnnotation = $NewSnippet.AppendChild($NewAnnotation)
                                    $NewSnippet.annotation = spanify $Spell.school
                                } # if ($AssignTag)
                                #>
                            } # "School"

                            "Subschool" {
                                $NewSnippet.setattribute("type","Tag_Standard")

                                # This code is in lieu of tags working properly.
                                $NewAnnotation = $NewSnippet.OwnerDocument.CreateElement("annotation",$NewSnippet.NamespaceURI)
                                $NewAnnotation = $NewSnippet.AppendChild($NewAnnotation)
                                $NewSnippet.annotation = $Spell.subschool

                                <# The code below is deactivated until tags are working properly.
                                $SnippetTags = $structure.domain_global | Where-Object {$_.name -eq "Spell Subschool"}
                                $AssignTag = ($SnippetTags.GetElementsByTagName("*") | Where-Object {$_.tag_id}) | Where-Object {$_.name -eq $Spell.school}
                                if ($AssignTag) {
                                    $NewTag = $NewSnippet.OwnerDocument.CreateElement("tag_assign",$NewSnippet.NameSpaceURI)
                                    $NewTag.setattribute("tag_id",$AssignTag.tag_id)
                                    $NewTag.setattribute("type","Indirect")
                                    $NewTag = $NewSnippet.AppendChild($NewTag)
                                } elseif (($Spell.subschool.contains(",")) -or ($Spell.subschool.contains(" or ")) -or ($Spell.subschool.contains("see text"))) {
                                    $AssignTag = ($SnippetTags.GetElementsByTagName("*") | Where-Object {$_.tag_id}) | Where-Object {$_.name -eq "see text or annotation"}
                                    $NewAnnotation = $NewSnippet.OwnerDocument.CreateElement("annotation",$NewSnippet.NamespaceURI)
                                    $NewAnnotation = $NewSnippet.AppendChild($NewAnnotation)
                                    $NewSnippet.annotation = spanify $Spell.subschool
                                } # if ($AssignTag)
                                #>
                            } # "Subschool"
                            
                            "Descriptor(s)" {
                                $NewSnippet.setattribute("type","Tag_Standard")

                                # This code is in lieu of tags working properly.
                                $NewAnnotation = $NewSnippet.OwnerDocument.CreateElement("annotation",$NewSnippet.NamespaceURI)
                                $NewAnnotation = $NewSnippet.AppendChild($NewAnnotation)
                                $NewSnippet.annotation = $Spell.Descriptor

                                <# The code below is deactivated until tags are working properly.
                                $SnippetTags = $structure.domain_global | Where-Object {$_.name -eq "Spell Descriptor"}
                                $Descriptor = $Spell.Descriptor
                                if (($Descriptor) -and (-not $Descriptor.contains(" or ")) -and (-not $Descriptor.contains("see text")) -and (-not $Descriptor.contains("variable"))) {
                                    $DescriptorList = $Descriptor.split(",")
                                    foreach ($Descriptor in $DescriptorList) {
                                        $Descriptor = $Descriptor.Trim()
                                        $AssignTag = ($SnippetTags.GetElementsByTagName("*") | Where-Object {$_.tag_id}) | Where-Object {$_.name -eq $Descriptor}
                                        if (-not $AssignTag.tag_id) {$Descriptor >> 'c:\Users\Jim\Desktop\RWImport\tags.txt'}
                                        $NewTag = $NewSnippet.OwnerDocument.CreateElement("tag_assign",$NewSnippet.NameSpaceURI)
                                        $NewTag.setattribute("tag_id",$AssignTag.tag_id)
                                        $NewTag.setattribute("type","Indirect")
                                        $NewTag = $NewSnippet.AppendChild($NewTag)
                                    } # foreach ($Descriptor in $DescriptorList)
                                } elseif (($Descriptor) -and (($Descriptor.contains(" or ")) -or ($Descriptor.contains("see text")) -or ($Descriptor.contains("variable")))) {
                                    $AssignTag = ($SnippetTags.GetElementsByTagName("*") | Where-Object {$_.tag_id}) | Where-Object {$_.name -eq "see text or annotation"}
                                    $NewAnnotation = $NewSnippet.OwnerDocument.CreateElement("annotation",$NewSnippet.NamespaceURI)
                                    $NewAnnotation = $NewSnippet.AppendChild($NewAnnotation)
                                    $NewSnippet.annotation = spanify $Spell.Descriptor
                                } # if ($Descriptor)
                                #>
                            } # "Descriptor(s)"
                            
                            "Level" {
                                $NewSnippet.setattribute("type","Tag_Standard")

                                # This code is in lieu of tags working properly.
                                $NewAnnotation = $NewSnippet.OwnerDocument.CreateElement("annotation",$NewSnippet.NamespaceURI)
                                $NewAnnotation = $NewSnippet.AppendChild($NewAnnotation)
                                $NewSnippet.annotation = $Spell.Spell_Level

                                <# The code below is deactivated until tags are working properly.
                                $SnippetTags = $structure.domain_global | Where-Object {$_.name -eq "Spell Level"}
                                $Level = $Spell.Spell_Level
                                if ($Level) {
                                    $LevelList = $Level.split(",")
                                    foreach ($LevelTag in $LevelList) {
                                        $LevelTag = $LevelTag.Trim()
                                        $AssignTag = ($SnippetTags.GetElementsByTagName("*") | Where-Object {$_.tag_id}) | Where-Object {$_.name -eq $LevelTag}
                                        if ($AssignTag) {
                                            $NewTag = $NewSnippet.OwnerDocument.CreateElement("tag_assign",$NewSnippet.NameSpaceURI)
                                            $NewTag.setattribute("tag_id",$AssignTag.tag_id)
                                            $NewTag.setattribute("type","Indirect")
                                            $NewTag = $NewSnippet.AppendChild($NewTag)
                                        } # if ($AssignTag)
                                    } # foreach ($Level in $LevelList)
                                } # if ($Level)
                                #>
                            } # "Level"
                    
                            "Casting Time" {
                                $NewSnippet.setattribute("type","Labeled_Text")
                                $NewContent = $NewSnippet.OwnerDocument.CreateElement("contents",$NewSnippet.NameSpaceURI)
                                $NewContent = $NewSnippet.AppendChild($NewContent)
                                $NewSnippet.contents = $Spell.casting_time
                            } # "Casting Time"

                            "Components" {
                                $NewSnippet.setattribute("type","Labeled_Text")
                                $NewContent = $NewSnippet.OwnerDocument.CreateElement("contents",$NewSnippet.NameSpaceURI)
                                $NewContent = $NewSnippet.AppendChild($NewContent)
                                $NewSnippet.contents = $Spell.components
                            } # "Components"

                            "Range" {
                                $NewSnippet.setattribute("type","Tag_Standard")

                                # This code is in lieu of tags working properly.
                                $NewAnnotation = $NewSnippet.OwnerDocument.CreateElement("annotation",$NewSnippet.NamespaceURI)
                                $NewAnnotation = $NewSnippet.AppendChild($NewAnnotation)
                                $NewSnippet.annotation = $Spell.Range

                                <# The code below is deactivated until tags are working properly.
                                $SnippetTags = $structure.domain_global | Where-Object {$_.name -eq "Spell Range"}
                                $AssignTag = ($SnippetTags.GetElementsByTagName("*") | Where-Object {$_.tag_id}) | Where-Object {$_.name -eq $Spell.Range}
                                if ($AssignTag) {
                                    $NewTag = $NewSnippet.OwnerDocument.CreateElement("tag_assign",$NewSnippet.NameSpaceURI)
                                    $NewTag.setattribute("tag_id",$AssignTag.tag_id)
                                    $NewTag.setattribute("type","Indirect")
                                    $NewTag = $NewSnippet.AppendChild($NewTag)
                                } elseif (($Spell.Range.contains(",")) -or ($Spell.Range.contains(" or ")) -or ($Spell.Range.contains("see text"))) {
                                    $AssignTag = ($SnippetTags.GetElementsByTagName("*") | Where-Object {$_.tag_id}) | Where-Object {$_.name -eq "see text or annotation"}
                                    $NewAnnotation = $NewSnippet.OwnerDocument.CreateElement("annotation",$NewSnippet.NamespaceURI)
                                    $NewAnnotation = $NewSnippet.AppendChild($NewAnnotation)
                                    $NewSnippet.annotation = spanify $Spell.Range
                                } # if ($AssignTag)
                                #>

                            } # "Range"
                            
                            "Effect" {
                                $NewSnippet.setattribute("type","Labeled_Text")
                                $NewContent = $NewSnippet.OwnerDocument.CreateElement("contents",$NewSnippet.NameSpaceURI)
                                $NewContent = $NewSnippet.AppendChild($NewContent)
                                $NewSnippet.contents = $Spell.effect
                            } # "Effect"

                            "Target" {
                                $NewSnippet.setattribute("type","Labeled_Text")
                                $NewContent = $NewSnippet.OwnerDocument.CreateElement("contents",$NewSnippet.NameSpaceURI)
                                $NewContent = $NewSnippet.AppendChild($NewContent)
                                $NewSnippet.contents = $Spell.targets
                            } # "Target"

                            "Area" {
                                $NewSnippet.setattribute("type","Labeled_Text")
                                $NewContent = $NewSnippet.OwnerDocument.CreateElement("contents",$NewSnippet.NameSpaceURI)
                                $NewContent = $NewSnippet.AppendChild($NewContent)
                                $NewSnippet.contents = $Spell.area
                            } # "Area"

                            "Duration" {
                                $NewSnippet.setattribute("type","Labeled_Text")
                                $NewContent = $NewSnippet.OwnerDocument.CreateElement("contents",$NewSnippet.NameSpaceURI)
                                $NewContent = $NewSnippet.AppendChild($NewContent)
                                $NewSnippet.contents = $Spell.duration
                            } # "Duration"

                            "Saving Throw" {
                                $NewSnippet.setattribute("type","Labeled_Text")
                                $NewContent = $NewSnippet.OwnerDocument.CreateElement("contents",$NewSnippet.NameSpaceURI)
                                $NewContent = $NewSnippet.AppendChild($NewContent)
                                $NewSnippet.contents = $Spell.saving_throw
                            } # "Saving Throw"

                            "Spell Resistance" {
                                $NewSnippet.setattribute("type","Labeled_Text")
                                $NewContent = $NewSnippet.OwnerDocument.CreateElement("contents",$NewSnippet.NameSpaceURI)
                                $NewContent = $NewSnippet.AppendChild($NewContent)
                                $NewSnippet.contents = $Spell.spell_resistance
                            } # "Spell Resistance"
                        } # Switch ($Facet.name)
                    } # foreach ($Facet in $FacetList)
                } # "Overview"

                "Description" {
                    $NewSnippet = $NewSection.OwnerDocument.CreateElement("snippet",$NewSection.NameSpaceURI)
                    $NewSnippet.setattribute("type","Multi_Line")
                    $NewSnippet = $NewSection.AppendChild($NewSnippet)
                    $NewContent = $NewSnippet.OwnerDocument.CreateElement("contents",$NewSnippet.NameSpaceURI)
                    $NewContent = $NewSnippet.AppendChild($NewContent)
                    
                    $NewSnippet.contents = $Spell.description_formatted
                    # $NewSnippet.contents = spanify $Spell.description_formatted
                    # $NewSnippet.contents = $Spell.description
                } # "Description"

                "Mythic" {
                    $NewSnippet = $NewSection.OwnerDocument.CreateElement("snippet",$NewSection.NameSpaceURI)
                    $NewSnippet.setattribute("type","Multi_Line")
                    $NewSnippet = $NewSection.AppendChild($NewSnippet)
                } # "Mythic"
            } # Switch (Section.name)
        } # foreach ($Section in $SectionList)

    } # for ($SpellCount = 0; $SpellCount -le 0; $SpellCount++)
} # Function ImportSpells($structure,$SpellContainer,$SpellCategory,$SpellFile)

# Main {

#############################################################################
#############   E D I T   Y O U R   F I L N A M E S   H E R E   #############
#############################################################################
####
<##> $Structure_File = "Pathfinder_Structure_Augmented.rwexport"
<##> $Spell_Spreadsheet = "spell_full - Updated 29Jan2017 - sanitized.csv"
<##> $New_XML_File = "Pathfinder_Spells.rwexport"
####
#############################################################################

# These are global counters, across all functions
# to give each topic and each alias it's own, unique,
# sequential topic_id or alias_id. 
$Script:topic_id = 1
$Script:alias_id = 1

# Load the structure file, preserving the existing white space configuration.
$RWExportData = New-Object xml
$RWExportData.PreserveWhitespace = $true
$RWExportData.Load($Structure_File)

# Get a list of all elements with a category_id value.
# Loading the categories this way gives us both
# global and user categories in one shot.
$structure = $RWExportData.export.structure
$CategoryList = $structure.GetElementsByTagName("*") | Where-Object {$_.category_id}

# If the contents element is empty, then loading it directly by name will make
# Powershell think it's an empty string instead of an XML object.
# This method forces Powershell to load it as an XML object, even if it's empty.
$contents = $RWExportData.export.ChildNodes | Where-Object {$_.name -eq "contents"}

# Create a General Abilities Article as a container for spells.
$FindCategory = $CategoryList | Where-Object {$_.name -eq "General Abilities Article"}
$ContainerCategory = New-Object objIDMap
$ContainerCategory.name = "Spells"
$ContainerCategory.category_id = $FindCategory.category_id
$ContainerCategory.alias = "Spell"
$SpellContainer = CreateContainer $structure $contents $ContainerCategory

# Get the spell category fromt he category list we created earlier,
# then invoke the ImportSpells function.
$SpellCategory = $CategoryList | Where-Object {$_.name -eq "Spell"}

# Create subcontainers A-Z under the Spell container.
$SpellFile = $Spell_Spreadsheet
$SpellList = Import-Csv $SpellFile
$Alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
for ($Letter = 0; $Letter -le 25; $Letter++) {
    $FindCategory = $CategoryList | Where-Object {$_.name -eq "General Abilities Article"}
    $ContainerCategory = New-Object objIDMap
    $ContainerCategory.name = "Spells - " + $Alphabet[$Letter]
    $ContainerCategory.category_id = $FindCategory.category_id
    $LetterContainer = CreateContainer $structure $SpellContainer $ContainerCategory

    # Get all the spells that start with the current letter, and import them into
    # the current letter's container.
    $LetterSpells = $SpellList | Where-object {$_.name.startswith($Alphabet[$Letter])}
    ImportSpells $structure $LetterContainer $SpellCategory $LetterSpells
} # for ($Letter = 0; $Letter -le 25; $Letter++)


# Save the modified XML data to a new file.
$RWExportData.Save($New_XML_File)
