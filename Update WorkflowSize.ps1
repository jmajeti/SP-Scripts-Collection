 get-WFServiceConfiguration -ServiceUri http://sp-app-s1:12291/ -Name WorkflowServiceMaxArgumentsPerActivity
The default value is 50
Set-WFServiceConfiguration -ServiceUri https://WORKFLOWSERVER.FQDN:12290 -Name WorkflowServiceMaxArgumentsPerActivity -Value 100
Set-WFServiceConfiguration -ServiceUri: http://sp-app-s1:12291/ -Name: WorkflowServiceMaxWorkflowXamlSizeInBytes -Value: 104857600
Set-WFServiceConfiguration -ServiceUri http://sp-app-s1:12291/ -Name 
WorkflowServiceMaxInstanceSizeKB -Value 102400
$webapp = Get-SPWebApplication -identity https://sharepoint.hrsa.gov
$webapp.UpdateWorkflowConfigurationSettings()