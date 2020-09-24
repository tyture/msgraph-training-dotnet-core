using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace GraphTutorial
{
    public class GraphHelper
    {
        private static GraphServiceClient graphClient;
        public static void Initialize(IAuthenticationProvider authProvider)
        {
            graphClient = new GraphServiceClient(authProvider);
        }

        public static async Task<User> GetMeAsync()
        {
            try
            {
                // GET /me
                return await graphClient.Me.Request().GetAsync();
            }
            catch (ServiceException ex)
            {
                Console.WriteLine($"Error getting signed-in user: {ex.Message}");
                return null;
            }
        }

        // <GetEventsSnippet>
        public static async Task<IEnumerable<Event>> GetEventsAsync()
        {
            try
            {
                // GET /me/events
                var resultPage = await graphClient.Me.Events.Request()
                    // Only return the fields used by the application
                    .Select(e => new {
                      e.Subject,
                      e.Organizer,
                      e.Start,
                      e.End
                    })
                    // Sort results by when they were created, newest first
                    .OrderBy("createdDateTime DESC")
                    .GetAsync();

                return resultPage.CurrentPage;
            }
            catch (ServiceException ex)
            {
                Console.WriteLine($"Error getting events: {ex.Message}");
                return null;
            }
        }
        // </GetEventsSnippet>

        public static async Task<ISearchQueryCollectionPage> GetItem(string connectionId, string query)
        {
            var requests = new List<SearchRequestObject>()
            {
                new SearchRequestObject
                {
                    EntityTypes = new List<EntityType>()
                    {
                        EntityType.ExternalItem
                    },
                    ContentSources = new List<String>()
                    {
                        $"/external/connections/{connectionId}"
                    },
                    Query = new SearchQuery
                    {
                        Query_string = new SearchQueryString
                        {
                            Query = query
                        }
                    },
                    From = 0,
                    Size = 5,
                    Stored_fields = new List<String>()
                    {
                        "partNumber",
                        "name",
                        "description",
                        "price",
                        "inventory",
                        "appliances@odata.type",
                        "appliances",
                    }
                }
            };

            return await graphClient.Search
                .Query(requests)
                .Request()
                .PostAsync();
        }
    }
}