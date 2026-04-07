import { list, put } from "@vercel/blob";

type AnnotationMap = Record<number, unknown[]>;

type PersistedProjectPayload = {
  id: string;
  name: string;
  fileName: string;
  annotations: AnnotationMap;
  calibration: number;
  pageTexts: Record<number, string>;
  pageSheetNumbers?: Record<number, string>;
  pdfBlobUrl?: string;
  layers: unknown[];
  users: unknown[];
  updatedAt: string;
};

type CloudProjectSummary = {
  id: string;
  name: string;
  fileName: string;
  updatedAt: string;
  pdfBlobUrl?: string;
  projectDataUrl: string;
};

function sanitizePathSegment(value: string): string {
  return value
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9._-]+/g, "-")
    .replace(/-{2,}/g, "-")
    .replace(/^-+|-+$/g, "") || "project";
}

function toSummary(project: PersistedProjectPayload, projectDataUrl: string): CloudProjectSummary {
  return {
    id: project.id,
    name: project.name,
    fileName: project.fileName,
    updatedAt: project.updatedAt,
    pdfBlobUrl: project.pdfBlobUrl,
    projectDataUrl,
  };
}

async function listProjects() {
  const { blobs } = await list({ prefix: "projects/" });
  const projectBlobs = blobs.filter((blob) => blob.pathname.endsWith("/project.json"));
  const projects = await Promise.all(
    projectBlobs.map(async (blob) => {
      const response = await fetch(blob.url, { cache: "no-store" });
      if (!response.ok) return null;
      const project = (await response.json()) as PersistedProjectPayload;
      return toSummary(project, blob.url);
    })
  );

  return projects
    .filter((project): project is CloudProjectSummary => Boolean(project))
    .sort((a, b) => String(b.updatedAt).localeCompare(String(a.updatedAt)));
}

export default async function handler(request: Request): Promise<Response> {
  if (request.method === "GET") {
    try {
      const projects = await listProjects();
      return Response.json({ projects });
    } catch (error) {
      const message = error instanceof Error ? error.message : "Unable to list shared projects";
      return Response.json({ error: message }, { status: 500 });
    }
  }

  if (request.method === "POST") {
    try {
      const project = (await request.json()) as PersistedProjectPayload;
      if (!project?.id || !project?.name) {
        return Response.json({ error: "Project id and name are required" }, { status: 400 });
      }

      const payload: PersistedProjectPayload = {
        ...project,
        updatedAt: new Date().toISOString(),
      };

      const blob = await put(
        `projects/${sanitizePathSegment(project.id)}/project.json`,
        JSON.stringify(payload, null, 2),
        {
          access: "public",
          contentType: "application/json",
          allowOverwrite: true,
          addRandomSuffix: false,
        }
      );

      return Response.json({ ok: true, project: toSummary(payload, blob.url) });
    } catch (error) {
      const message = error instanceof Error ? error.message : "Unable to save shared project";
      return Response.json({ error: message }, { status: 500 });
    }
  }

  return new Response("Method not allowed", { status: 405 });
}
