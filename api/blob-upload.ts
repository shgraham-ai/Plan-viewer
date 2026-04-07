import { handleUpload, type HandleUploadBody } from "@vercel/blob/client";

export const config = {
  api: {
    bodyParser: false,
  },
};

export default async function handler(request: Request): Promise<Response> {
  if (request.method !== "POST") {
    return new Response("Method not allowed", { status: 405 });
  }

  const body = (await request.json()) as HandleUploadBody;

  try {
    const jsonResponse = await handleUpload({
      body,
      request,
      onBeforeGenerateToken: async (pathname) => ({
        allowedContentTypes: ["application/pdf"],
        addRandomSuffix: false,
        allowOverwrite: true,
        tokenPayload: JSON.stringify({
          pathname,
          createdAt: new Date().toISOString(),
        }),
      }),
      onUploadCompleted: async () => {
        // No-op for now. Project metadata is stored separately in /api/cloud-projects.
      },
    });

    return Response.json(jsonResponse);
  } catch (error) {
    const message = error instanceof Error ? error.message : "Blob upload failed";
    return Response.json({ error: message }, { status: 400 });
  }
}
