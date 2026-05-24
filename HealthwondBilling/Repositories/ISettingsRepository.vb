Imports HealthwondBilling.Models

Namespace Repositories

    Public Interface ISettingsRepository
        Function GetProfile() As AppSettingsProfile
        Sub SaveProfile(profile As AppSettingsProfile)
    End Interface

End Namespace
