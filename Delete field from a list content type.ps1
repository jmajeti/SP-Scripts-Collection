$web = Get-SPWeb http://localhost
$list = $web.Lists["MyList"]
$listContentTypes = $list.ContentTypes
$listContentType = $listContentTypes["My List Content Type"]
$listFieldLink = $listContentType.FieldLinks["MyListFieldInternalName"]
$listContentType.FieldLinks.Delete("MyListFieldInternalName")
$listContentType.Update()