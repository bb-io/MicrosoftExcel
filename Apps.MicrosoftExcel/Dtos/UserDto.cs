﻿using Blackbird.Applications.Sdk.Common;

namespace Apps.MicrosoftExcel.Dtos;

public class UserDto
{
    [Display("User ID")]
    public string Id { get; set; }
    
    public string Email { get; set; }
    
    [Display("Display name")]
    public string DisplayName { get; set; }
}