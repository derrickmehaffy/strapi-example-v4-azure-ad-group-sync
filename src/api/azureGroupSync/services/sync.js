const { Client } = require("@microsoft/microsoft-graph-client");
const { env } = require("@strapi/utils");

module.exports = ({ strapi }) => ({
  async syncUserGroups(event) {
    // initialize graph client auth
    let graphAuth;

    // build params for auth request
    const params = new URLSearchParams();
    params.append("client_id", env("MICROSOFT_CLIENT_ID", ""));
    params.append("client_secret", env("MICROSOFT_CLIENT_SECRET", ""));
    params.append("grant_type", "client_credentials");
    params.append("resource", "https://graph.microsoft.com");

    // make auth request
    try {
      authRequest = await strapi.fetch(
        "https://login.microsoftonline.com/strapitest.onmicrosoft.com/oauth2/token",
        {
          method: "POST",
          body: params,
        }
      );

      graphAuth = await authRequest.json();
    } catch (error) {
      console.log(error);
    }

    // initialize graph client using beta api
    const graphClient = Client.init({
      defaultVersion: "beta",
      authProvider: (done) => {
        done(null, graphAuth.access_token);
      },
    });

    // Validate user is logging in via Azure AD
    if (graphAuth.access_token && event.provider === "azure_ad_oauth2") {
      // Get user from graph api
      const user = await graphClient
        .api("/users")
        .filter(`otherMails/any(id:id eq '${event.user.email}')`)
        .get();

      if (user.value.length > 0) {
        // Get groups from graph api
        const groups = await graphClient
          .api(`/users/${user.value[0].id}/memberOf`)
          .get();

        // Map out the role names from the groups
        const roles = groups.value.map((group) => {
          if (group.displayName) {
            return group.displayName;
          }
        });

        // Find the roles in Strapi
        const strapiRoles = await strapi.entityService.findMany("admin::role", {
          filters: {
            name: {
              $in: roles.filter(Boolean),
            },
          },
        });

        // Map out the strapi roles and clean them up
        let userRoles = strapiRoles.map((role) => role.id);
        userRoles.filter(Boolean);

        // If no roles are found, default to super admin (prob should be different)
        if (!userRoles) {
          userRoles = [1];
        }

        // Update the user with the roles
        await strapi.entityService.update("admin::user", event.user.id, {
          data: {
            roles: userRoles,
          },
        });
      }
    }
  },
});
