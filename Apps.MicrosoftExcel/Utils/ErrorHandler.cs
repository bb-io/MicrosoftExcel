﻿using Blackbird.Applications.Sdk.Common.Exceptions;

namespace Apps.MicrosoftExcel.Utils
{
    public static class ErrorHandler
    {
        public static async Task ExecuteWithErrorHandlingAsync(Func<Task> action)
        {
            try
            {
                await action();
            }

            catch (Exception ex)
            {
                throw new PluginApplicationException(ex.Message);
            }
        }

        public static async Task<T> ExecuteWithErrorHandlingAsync<T>(Func<Task<T>> action)
        {
            try
            {
                return await action();
            }
            catch (Exception ex)
            {
                throw new PluginApplicationException(ex.Message);
            }
        }
    }
}
