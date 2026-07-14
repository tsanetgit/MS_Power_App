# Copilot Instructions

## Project Guidelines
- When reading shared/system configuration entities (e.g., integration settings, environment variables) in Dataverse plugins, use IOrganizationServiceFactory.CreateOrganizationService impersonating the resolved non-interactive SYSTEM user (accessmode=3, isdisabled=false) rather than the calling user's context, since the calling user may lack read privileges on those entities.