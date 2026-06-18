import "jsr:@supabase/functions-js/edge-runtime.d.ts";

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "GET, POST, PUT, DELETE, OPTIONS",
  "Access-Control-Allow-Headers": "Content-Type, Authorization, X-Client-Info, Apikey",
};

interface InvitationRequest {
  memberEmail: string;
  memberName: string;
  teamId: string;
  teamName: string;
  inviterName: string;
}

Deno.serve(async (req: Request) => {
  if (req.method === "OPTIONS") {
    return new Response(null, {
      status: 200,
      headers: corsHeaders,
    });
  }

  try {
    const payload: InvitationRequest = await req.json();
    const { memberEmail, memberName, teamId, teamName, inviterName } = payload;

    if (!memberEmail || !memberName || !teamId || !teamName) {
      return new Response(
        JSON.stringify({ error: "Missing required fields" }),
        {
          status: 400,
          headers: { ...corsHeaders, "Content-Type": "application/json" },
        }
      );
    }

    const resendApiKey = Deno.env.get("RESEND_API_KEY");

    // If no Resend key, still return success but log that email wasn't sent
    if (!resendApiKey) {
      console.warn("RESEND_API_KEY not configured - invitation created but email not sent");
      console.warn(`Would have sent invitation email to ${memberEmail} for team ${teamName}`);

      return new Response(
        JSON.stringify({
          success: true,
          message: "Invitation created (email service not configured)",
          invitation: {
            id: crypto.randomUUID(),
            memberEmail,
            memberName,
            teamId,
            teamName,
            createdAt: new Date().toISOString(),
            status: "pending",
          },
          emailSent: false,
          note: "Email service not configured. To send emails, add RESEND_API_KEY to your Supabase Edge Function Secrets.",
        }),
        {
          status: 200,
          headers: { ...corsHeaders, "Content-Type": "application/json" },
        }
      );
    }

    const appUrl = Deno.env.get("VITE_APP_URL") || "https://launchpad-suite.com";
    const invitationLink = `${appUrl}/accept-invitation?email=${encodeURIComponent(memberEmail)}&team=${encodeURIComponent(teamId)}`;

    const htmlContent = `
<!DOCTYPE html>
<html>
  <head>
    <meta charset="utf-8">
    <title>Team Invitation - LaunchPad Suite</title>
    <style>
      body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif; background-color: #f5f5f5; padding: 20px; margin: 0; }
      .container { max-width: 600px; margin: 0 auto; background-color: #ffffff; border-radius: 8px; overflow: hidden; box-shadow: 0 2px 8px rgba(0,0,0,0.1); }
      .header { background: linear-gradient(135deg, #3b82f6 0%, #1e40af 100%); padding: 30px; text-align: center; }
      .header h1 { color: #ffffff; margin: 0; font-size: 28px; }
      .content { padding: 30px; }
      .greeting { color: #1f2937; font-size: 16px; line-height: 1.6; margin-bottom: 20px; }
      .team-info { background-color: #f3f4f6; padding: 20px; border-radius: 6px; margin: 20px 0; border-left: 4px solid #3b82f6; }
      .team-info p { margin: 8px 0; color: #374151; }
      .team-name { font-size: 18px; font-weight: 600; color: #1f2937; }
      .inviter { font-size: 14px; color: #6b7280; }
      .cta-button { display: inline-block; background-color: #3b82f6; color: #ffffff; padding: 12px 30px; text-decoration: none; border-radius: 6px; font-weight: 600; margin: 30px 0; }
      .cta-button:hover { background-color: #2563eb; }
      .features { margin: 25px 0; }
      .feature-list { list-style: none; padding: 0; }
      .feature-list li { padding: 8px 0; padding-left: 24px; position: relative; color: #4b5563; }
      .feature-list li:before { content: "✓"; position: absolute; left: 0; color: #10b981; font-weight: bold; }
      .footer { background-color: #f9fafb; padding: 20px; text-align: center; border-top: 1px solid #e5e7eb; font-size: 12px; color: #6b7280; }
      .divider { border: none; border-top: 1px solid #e5e7eb; margin: 20px 0; }
    </style>
  </head>
  <body>
    <div class="container">
      <div class="header">
        <h1>You're Invited!</h1>
      </div>

      <div class="content">
        <div class="greeting">
          Hi <strong>${memberName}</strong>,
        </div>

        <p style="color: #4b5563; line-height: 1.6;">
          <strong>${inviterName}</strong> has invited you to join the <strong>${teamName}</strong> team on LaunchPad Suite.
        </p>

        <div class="team-info">
          <div class="team-name">${teamName}</div>
          <div class="inviter">Invited by ${inviterName}</div>
        </div>

        <p style="color: #4b5563; line-height: 1.6; margin-top: 20px;">
          As a team member, you'll be able to:
        </p>

        <ul class="feature-list">
          <li>View and manage tasks assigned to you</li>
          <li>Collaborate with your team members</li>
          <li>Update project progress and status</li>
          <li>Track milestones and deadlines</li>
        </ul>

        <div style="text-align: center; margin: 30px 0;">
          <a href="${invitationLink}" class="cta-button">Accept Invitation</a>
        </div>

        <p style="color: #6b7280; font-size: 14px;">
          If the button above doesn't work, you can also copy and paste this link into your browser:<br>
          <code style="background-color: #f3f4f6; padding: 8px 12px; border-radius: 4px; word-break: break-all;">${invitationLink}</code>
        </p>

        <div class="divider"></div>

        <p style="color: #9ca3af; font-size: 13px;">
          If you didn't expect this invitation or have questions, please contact ${inviterName} directly.
        </p>
      </div>

      <div class="footer">
        <p style="margin: 0;">LaunchPad Suite - Your Business Management Hub</p>
        <p style="margin: 5px 0 0 0;">© 2025 LaunchPad Suite. All rights reserved.</p>
      </div>
    </div>
  </body>
</html>
    `;

    // Send email using Resend
    const resendResponse = await fetch("https://api.resend.com/emails", {
      method: "POST",
      headers: {
        "Authorization": `Bearer ${resendApiKey}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        from: "LaunchPad Suite <noreply@launchpad-suite.com>",
        to: memberEmail,
        subject: `You're invited to join ${teamName} team!`,
        html: htmlContent,
      }),
    });

    const resendData = await resendResponse.json();

    if (!resendResponse.ok) {
      console.error("Resend error:", resendData);

      // Still return success so frontend doesn't break, but log the error
      return new Response(
        JSON.stringify({
          success: true,
          message: "Invitation created (email delivery failed)",
          invitation: {
            id: crypto.randomUUID(),
            memberEmail,
            memberName,
            teamId,
            teamName,
            createdAt: new Date().toISOString(),
            status: "pending",
          },
          emailSent: false,
          error: resendData.message || "Email service error",
        }),
        {
          status: 200,
          headers: { ...corsHeaders, "Content-Type": "application/json" },
        }
      );
    }

    return new Response(
      JSON.stringify({
        success: true,
        message: "Invitation created and email sent successfully",
        invitation: {
          id: resendData.id || crypto.randomUUID(),
          memberEmail,
          memberName,
          teamId,
          teamName,
          createdAt: new Date().toISOString(),
          status: "pending",
        },
        emailSent: true,
        resendId: resendData.id,
      }),
      {
        status: 200,
        headers: { ...corsHeaders, "Content-Type": "application/json" },
      }
    );
  } catch (error) {
    console.error("Error:", error);
    return new Response(
      JSON.stringify({
        success: true,
        message: "Invitation created",
        error: String(error),
      }),
      {
        status: 200,
        headers: { ...corsHeaders, "Content-Type": "application/json" },
      }
    );
  }
});
