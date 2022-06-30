using System;
using System.Collections.Generic;
using Microsoft.Identity.Client;
using Microsoft.Graph;
using Microsoft.Extensions.Configuration;
using Helpers;

namespace graphconsoleapp
{
    public class Program
    {
        public static async Task Main(string[] args)
        {
            Console.WriteLine("Hello World!");

            var config = LoadAppSettings();
            if (config == null)
            {
            Console.WriteLine("Invalid appsettings.json file.");
            return;
            }

            var client = GetAuthenticatedGraphClient(config);

            var graphRequest = client.Users.Request();

            var results = graphRequest.GetAsync().Result;
            foreach(var user in results)
            {
            Console.WriteLine(user.Id + ": " + user.DisplayName + " <" + user.Mail + ">");
            }

            Console.WriteLine("\nGraph Request:");
            Console.WriteLine(graphRequest.GetHttpRequestMessage().RequestUri);

            await CreateMeeting(client);
        }

        private static GraphServiceClient? _graphClient;

        private static IConfigurationRoot? LoadAppSettings()
        {
            try
            {
                var config = new ConfigurationBuilder()
                                .SetBasePath(System.IO.Directory.GetCurrentDirectory())
                                .AddJsonFile("appsettings.json", false, true)
                                .Build();

                if (string.IsNullOrEmpty(config["applicationId"]) ||
                    string.IsNullOrEmpty(config["applicationSecret"]) ||
                    string.IsNullOrEmpty(config["redirectUri"]) ||
                    string.IsNullOrEmpty(config["tenantId"]))
                {
                return null;
                }

                return config;
            }
            catch (System.IO.FileNotFoundException)
            {
                return null;
            }
        }

        private static IAuthenticationProvider CreateAuthorizationProvider(IConfigurationRoot config)
        {
            var clientId = config["applicationId"];
            var clientSecret = config["applicationSecret"];
            var redirectUri = config["redirectUri"];
            var authority = $"https://login.microsoftonline.com/{config["tenantId"]}/v2.0";

            List<string> scopes = new List<string>();
            scopes.Add("https://graph.microsoft.com/.default");

            var cca = ConfidentialClientApplicationBuilder.Create(clientId)
                                                    .WithAuthority(authority)
                                                    .WithRedirectUri(redirectUri)
                                                    .WithClientSecret(clientSecret)
                                                    .Build();
            return new MsalAuthenticationProvider(cca, scopes.ToArray());
        }

        private static GraphServiceClient GetAuthenticatedGraphClient(IConfigurationRoot config)
        {
            var authenticationProvider = CreateAuthorizationProvider(config);
            _graphClient = new GraphServiceClient(authenticationProvider);
            return _graphClient;
        }

        async static Task CreateMeeting(GraphServiceClient client)
        {
            var @event = new Event
            {
                Subject = "Let's go for lunch",
                Body = new ItemBody
                {
                    ContentType = BodyType.Html,
                    Content = "Does noon work for you?"
                },
                Start = new DateTimeTimeZone
                {
                    DateTime = "2022-07-15T12:00:00",
                    TimeZone = "Pacific Standard Time"
                },
                End = new DateTimeTimeZone
                {
                    DateTime = "2022-07-15T14:00:00",
                    TimeZone = "Pacific Standard Time"
                },
                Location = new Location
                {
                    DisplayName = "Harry's Bar"
                },
                Attendees = new List<Attendee>()
                {
                    new Attendee
                    {
                        EmailAddress = new EmailAddress
                        {
                            Address = "samanthab@contoso.onmicrosoft.com",
                            Name = "Samantha Booth"
                        },
                        Type = AttendeeType.Required
                    }
                },
                AllowNewTimeProposals = true,
                IsOnlineMeeting = true,
                OnlineMeetingProvider = OnlineMeetingProviderType.TeamsForBusiness
            };
            Console.WriteLine("Hi");
            try{Console.WriteLine("try");}
            catch {Console.WriteLine("catch");}
            Console.WriteLine(@event);

            try {
                var results2 = await client.Users["ruka.sakurai@tl6j3.onmicrosoft.com"].Events.Request().Header("Prefer","outlook.timezone=\"Pacific Standard Time\"").AddAsync(@event);
                Console.WriteLine(results2);
            } catch (Exception ex) {
                Console.WriteLine(ex);
            }
            Console.WriteLine("Success");
        }
    }
}