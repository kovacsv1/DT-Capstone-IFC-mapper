import ifcopenshell

def add_property(ifc_file, element, pset_name, prop_name, value):
    pset = None
    for definition in element.IsDefinedBy:
        if (
            definition.RelatingPropertyDefinition.is_a("IfcPropertySet")
            and definition.RelatingPropertyDefinition.Name == pset_name
        ):
            pset = definition.RelatingPropertyDefinition
            break

    if pset is None:
        pset = ifc_file.createIfcPropertySet(
            GlobalId=ifcopenshell.guid.new(),
            OwnerHistory=element.OwnerHistory,
            Name=pset_name,
            HasProperties=[],
        )
        ifc_file.createIfcRelDefinesByProperties(
            GlobalId=ifcopenshell.guid.new(),
            OwnerHistory=element.OwnerHistory,
            RelatedObjects=[element],
            RelatingPropertyDefinition=pset,
        )

    for prop in pset.HasProperties:
        if prop.Name == prop_name:
            prop.NominalValue = ifc_file.create_entity("IfcText", str(value))
            return  # Overwrite

    prop = ifc_file.createIfcPropertySingleValue(
        prop_name,
        None,
        ifc_file.create_entity("IfcText", str(value)),
        None,
    )
    pset.HasProperties = list(pset.HasProperties) + [prop]