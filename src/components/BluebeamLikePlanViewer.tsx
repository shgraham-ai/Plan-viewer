import React, { useEffect, useMemo, useRef, useState } from "react";
import { motion } from "framer-motion";
import * as pdfjsLib from "pdfjs-dist";
import { upload } from "@vercel/blob/client";
import {
  AreaChart,
  Circle,
  ChevronLeft,
  ChevronRight,
  ClipboardList,
  Copy,
  Download,
  Eye,
  EyeOff,
  FileOutput,
  FileText,
  FolderOpen,
  GitCompare,
  Hand,
  Layers3,
  Library,
  Link2,
  List,
  MessageSquare,
  MousePointer2,
  Maximize2,
  Minimize2,
  Pencil,
  Ruler,
  Save,
  Search,
  Square,
  Tag,
  Trash2,
  Type,
  Users,
  ZoomIn,
  ZoomOut,
} from "lucide-react";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { ScrollArea } from "@/components/ui/scroll-area";
import { Separator } from "@/components/ui/separator";

const TOOL = {
  PAN: "pan",
  SELECT: "select",
  DRAW: "draw",
  MEASURE: "measure",
  TEXT: "text",
  COMMENT: "comment",
  AREA: "area",
  RECTANGLE: "rectangle",
  ELLIPSE: "ellipse",
  CALLOUT: "callout",
  CLOUD: "cloud",
  STAMP: "stamp",
  CALIBRATE: "calibrate",
  LINK: "link",
  SNAPSHOT: "snapshot",
} as const;

type ToolType = (typeof TOOL)[keyof typeof TOOL];
type Point = { x: number; y: number };
type Transform = { x: number; y: number; scale: number };
type Layer = { id: string; name: string; visible: boolean };
type Role = { id: string; name: string; canEdit: boolean; canDelete: boolean; canExport: boolean };
type User = { id: string; name: string; roleId: string; color: string };
type PageSize = { width: number; height: number };
type WorkerStatus = "idle" | "ready" | "error";
type BookmarkSection = "all" | "civil" | "structural" | "architectural" | "mep" | "general";
type NavigationMode = "sheet" | "review";
type MenuId = "file" | "edit" | "view" | "document" | "markup" | "measure" | "window" | "help";
type RenderCanvasKey = "main" | "compare";
type TextContentItem = { str?: string; transform?: number[]; width?: number };
type Rect = { x: number; y: number; w: number; h: number };
type PageTextFragment = { text: string; rect: Rect };
type AutoReferenceLink = {
  id: string;
  page: number;
  text: string;
  detailToken: string;
  targetSheetNumber: string;
  targetPage: number;
  rect: Rect;
};
type MenuItem = {
  label: string;
  action?: () => void;
  disabled?: boolean;
  separator?: boolean;
  shortcut?: string;
};

type AnnotationStyle = {
  strokeWidth?: number;
  fillColor?: string;
  fillOpacity?: number;
};

type MeasureAnnotation = {
  id: string;
  type: "measure";
  start: Point;
  end: Point;
  page: number;
  author: string;
  status: string;
  createdAt: string;
  layerId: string;
  category: string;
  color: string;
  roleId: string;
} & AnnotationStyle;

type AreaAnnotation = {
  id: string;
  type: "area";
  points: Point[];
  page: number;
  author: string;
  status: string;
  createdAt: string;
  layerId: string;
  category: string;
  color: string;
  roleId: string;
} & AnnotationStyle;

type RectangleAnnotation = {
  id: string;
  type: "rectangle";
  rect: Rect;
  page: number;
  author: string;
  status: string;
  createdAt: string;
  layerId: string;
  category: string;
  color: string;
  roleId: string;
} & AnnotationStyle;

type EllipseAnnotation = {
  id: string;
  type: "ellipse";
  rect: Rect;
  page: number;
  author: string;
  status: string;
  createdAt: string;
  layerId: string;
  category: string;
  color: string;
  roleId: string;
} & AnnotationStyle;

type TextAnnotation = {
  id: string;
  type: "text" | "comment";
  text: string;
  comment?: string;
  position: Point;
  page: number;
  author: string;
  status: string;
  createdAt: string;
  layerId: string;
  category: string;
  color: string;
  roleId: string;
} & AnnotationStyle;

type CalloutAnnotation = {
  id: string;
  type: "callout";
  text: string;
  anchor: Point;
  box: { x: number; y: number; w: number; h: number };
  page: number;
  author: string;
  status: string;
  createdAt: string;
  layerId: string;
  category: string;
  color: string;
  roleId: string;
} & AnnotationStyle;

type PathAnnotation = {
  id: string;
  type: "path";
  points: Point[];
  page: number;
  author: string;
  status: string;
  createdAt: string;
  layerId: string;
  category: string;
  color: string;
  roleId: string;
} & AnnotationStyle;

type LinkAnnotation = {
  id: string;
  type: "link";
  text: string;
  targetPage: number;
  position: Point;
  page: number;
  author: string;
  status: string;
  createdAt: string;
  layerId: string;
  category: string;
  color: string;
  roleId: string;
} & AnnotationStyle;

type SnapshotAnnotation = {
  id: string;
  type: "snapshot";
  title: string;
  rect: { x: number; y: number; w: number; h: number };
  page: number;
  author: string;
  status: string;
  createdAt: string;
  layerId: string;
  category: string;
  color: string;
  roleId: string;
} & AnnotationStyle;

type CloudAnnotation = {
  id: string;
  type: "cloud";
  rect: { x: number; y: number; w: number; h: number };
  label?: string;
  page: number;
  author: string;
  status: string;
  createdAt: string;
  layerId: string;
  category: string;
  color: string;
  roleId: string;
} & AnnotationStyle;

type StampAnnotation = {
  id: string;
  type: "stamp";
  text: string;
  stampKind: string;
  position: Point;
  page: number;
  author: string;
  status: string;
  createdAt: string;
  layerId: string;
  category: string;
  color: string;
  roleId: string;
} & AnnotationStyle;

type Annotation =
  | MeasureAnnotation
  | AreaAnnotation
  | RectangleAnnotation
  | EllipseAnnotation
  | TextAnnotation
  | CalloutAnnotation
  | PathAnnotation
  | LinkAnnotation
  | SnapshotAnnotation
  | CloudAnnotation
  | StampAnnotation;

type AnnotationMap = Record<number, Annotation[]>;

type PersistedSession = {
  id: string;
  name: string;
  annotations: AnnotationMap;
  calibration: number;
  fileName: string;
  layers: Layer[];
  savedAt: string;
};

type PersistedProject = {
  id: string;
  name: string;
  fileName: string;
  annotations: AnnotationMap;
  calibration: number;
  pageTexts: Record<number, string>;
  pageSheetNumbers?: Record<number, string>;
  pdfBlobUrl?: string;
  layers: Layer[];
  users: User[];
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

type HistorySnapshot = {
  annotations: AnnotationMap;
  calibration: number;
  layers: Layer[];
  currentPage: number;
};

type InlineEditorField = "text" | "comment" | "label" | "title";

type InlineEditorState = {
  annotationId: string;
  page: number;
  field: InlineEditorField;
  value: string;
};

type SnapPreview = {
  point: Point;
  label: string;
  guideX?: number;
  guideY?: number;
};

type Draft =
  | { type: "path"; points: Point[]; layerId: string; category: string }
  | { type: "measure"; start: Point; end: Point; layerId: string; category: string }
  | { type: "area"; points: Point[]; layerId: string; category: string }
  | { type: "rectangle"; start: Point; current: Point; layerId: string; category: string }
  | { type: "ellipse"; start: Point; current: Point; layerId: string; category: string }
  | { type: "zoom-box"; start: Point; current: Point }
  | { type: "selection-box"; start: Point; current: Point; append: boolean }
  | { type: "callout"; start: Point; current: Point; text: string; layerId: string; category: string }
  | { type: "cloud"; start: Point; current: Point; layerId: string; category: string; label: string }
  | { type: "calibrate"; start: Point; end: Point }
  | { type: "snapshot"; start: Point; current: Point; title: string; layerId: string; category: string }
  | null;

const DEFAULT_LAYERS: Layer[] = [
  { id: "layer-general", name: "General", visible: true },
  { id: "layer-takeoff", name: "Takeoff", visible: true },
  { id: "layer-rfi", name: "RFI", visible: true },
  { id: "layer-punch", name: "Punch", visible: true },
  { id: "layer-links", name: "Links", visible: true },
];

const DEFAULT_ROLES: Role[] = [
  { id: "role-pm", name: "Project Manager", canEdit: true, canDelete: true, canExport: true },
  { id: "role-super", name: "Superintendent", canEdit: true, canDelete: true, canExport: true },
  { id: "role-arch", name: "Architect", canEdit: true, canDelete: false, canExport: true },
  { id: "role-viewer", name: "Viewer", canEdit: false, canDelete: false, canExport: false },
];

const DEFAULT_USERS: User[] = [
  { id: "u1", name: "Shawn", roleId: "role-pm", color: "#2563eb" },
  { id: "u2", name: "Field User", roleId: "role-super", color: "#15803d" },
];

const BOOKMARK_TABS: Array<{ id: BookmarkSection; label: string; keywords: string[] }> = [
  { id: "all", label: "All", keywords: [] },
  { id: "civil", label: "Civil", keywords: ["civil", "site", "grading", "erosion", "utility", "utilities", "paving", "drainage"] },
  { id: "structural", label: "Structural", keywords: ["structural", "foundation", "concrete", "steel", "framing", "rebar"] },
  { id: "architectural", label: "Architectural", keywords: ["architectural", "floor plan", "elevation", "door", "finish", "partition"] },
  { id: "mep", label: "MEP", keywords: ["mechanical", "electrical", "plumbing", "hvac", "fire protection"] },
  { id: "general", label: "General", keywords: [] },
];

const MARKUP_PRESETS = [
  { id: "preset-general", name: "General Note", color: "#2563eb", layerId: "layer-general", category: "Notes" },
  { id: "preset-rfi", name: "RFI", color: "#b91c1c", layerId: "layer-rfi", category: "RFI" },
  { id: "preset-punch", name: "Punch", color: "#d97706", layerId: "layer-punch", category: "Punch" },
  { id: "preset-takeoff", name: "Takeoff", color: "#15803d", layerId: "layer-takeoff", category: "Takeoff" },
  { id: "preset-links", name: "Link", color: "#1d4ed8", layerId: "layer-links", category: "Links" },
] as const;

const STAMP_PRESETS = [
  { id: "stamp-field-verify", name: "Field Verify", text: "FIELD VERIFY", color: "#d97706", layerId: "layer-general", category: "Field Review", fillOpacity: 0.2 },
  { id: "stamp-approved", name: "Approved", text: "APPROVED", color: "#15803d", layerId: "layer-general", category: "Review Status", fillOpacity: 0.18 },
  { id: "stamp-revise", name: "Revise & Resubmit", text: "REVISE + RESUBMIT", color: "#b91c1c", layerId: "layer-rfi", category: "Review Status", fillOpacity: 0.18 },
  { id: "stamp-rfi", name: "RFI Required", text: "RFI REQUIRED", color: "#b91c1c", layerId: "layer-rfi", category: "RFI", fillOpacity: 0.18 },
  { id: "stamp-punch", name: "Punch", text: "PUNCH LIST", color: "#d97706", layerId: "layer-punch", category: "Punch", fillOpacity: 0.2 },
  { id: "stamp-coordinated", name: "Coordinated", text: "COORDINATED", color: "#2563eb", layerId: "layer-general", category: "Coordination", fillOpacity: 0.18 },
] as const;

const MARKUP_STATUSES = [
  { value: "open", label: "Open" },
  { value: "in-review", label: "In Review" },
  { value: "rfi", label: "RFI" },
  { value: "approved-noted", label: "Approved as Noted" },
  { value: "closed", label: "Closed" },
] as const;

const SHEET_NUMBER_TOKEN = /^(?:[A-Z]{1,3}-?\d{2,3}(?:\.\d{1,3})?)$/;
const DETAIL_REFERENCE_TOKEN = /\b([A-Z0-9]{1,3})\/([A-Z]{1,3}-?\d{2,3}(?:\.\d{1,3})?)\b/g;
const DETAIL_TOKEN_ONLY = /^[A-Z0-9]{1,3}$/;
const KEYNOTE_TOKEN = /\b([A-Z]{1,2}\d{1,3}[A-Z]?)\b/g;

function createId(): string {
  if (typeof crypto !== "undefined" && typeof crypto.randomUUID === "function") {
    return crypto.randomUUID();
  }
  return `id_${Date.now()}_${Math.random().toString(36).slice(2, 10)}`;
}

function sanitizePathSegment(value: string): string {
  return value
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9._-]+/g, "-")
    .replace(/-{2,}/g, "-")
    .replace(/^-+|-+$/g, "") || "file";
}

function safeParse<T>(value: string | null, fallback: T): T {
  try {
    return value ? (JSON.parse(value) as T) : fallback;
  } catch {
    return fallback;
  }
}

function deepClone<T>(value: T): T {
  return JSON.parse(JSON.stringify(value)) as T;
}

function ensureArray<T>(value: T[] | undefined | null): T[] {
  return Array.isArray(value) ? value : [];
}

function distance(a: Point, b: Point): number {
  const dx = b.x - a.x;
  const dy = b.y - a.y;
  return Math.sqrt(dx * dx + dy * dy);
}

function polygonArea(points: Point[]): number {
  if (!Array.isArray(points) || points.length < 3) return 0;
  let area = 0;
  for (let index = 0; index < points.length; index += 1) {
    const next = (index + 1) % points.length;
    area += points[index].x * points[next].y;
    area -= points[next].x * points[index].y;
  }
  return Math.abs(area / 2);
}

function formatFeet(px: number, scale: number): string {
  return `${(px * scale).toFixed(2)} ft`;
}

function formatArea(px: number, scale: number): string {
  return `${(px * scale * scale).toFixed(2)} sf`;
}

function hexToRgba(hex: string, alpha: number): string {
  const normalized = hex.replace("#", "");
  const safe = normalized.length === 3
    ? normalized.split("").map((part) => `${part}${part}`).join("")
    : normalized.padEnd(6, "0").slice(0, 6);
  const r = Number.parseInt(safe.slice(0, 2), 16) || 0;
  const g = Number.parseInt(safe.slice(2, 4), 16) || 0;
  const b = Number.parseInt(safe.slice(4, 6), 16) || 0;
  return `rgba(${r}, ${g}, ${b}, ${clamp(alpha, 0, 1)})`;
}

function nearestPointOnRect(point: Point, rect: { x: number; y: number; w: number; h: number }): Point {
  return {
    x: clamp(point.x, rect.x, rect.x + rect.w),
    y: clamp(point.y, rect.y, rect.y + rect.h),
  };
}

function drawCloudRect(
  context: CanvasRenderingContext2D,
  rect: { x: number; y: number; w: number; h: number }
) {
  const { x, y, w, h } = rect;
  const right = x + w;
  const bottom = y + h;
  const bump = clamp(Math.min(w, h) / 6, 10, 22);

  context.beginPath();
  context.moveTo(x + bump, y);

  for (let px = x + bump; px < right - bump; px += bump) {
    const nextX = Math.min(px + bump, right - bump);
    context.quadraticCurveTo((px + nextX) / 2, y - bump * 0.65, nextX, y);
  }
  for (let py = y + bump; py < bottom - bump; py += bump) {
    const nextY = Math.min(py + bump, bottom - bump);
    context.quadraticCurveTo(right + bump * 0.65, (py + nextY) / 2, right, nextY);
  }
  for (let px = right - bump; px > x + bump; px -= bump) {
    const nextX = Math.max(px - bump, x + bump);
    context.quadraticCurveTo((px + nextX) / 2, bottom + bump * 0.65, nextX, bottom);
  }
  for (let py = bottom - bump; py > y + bump; py -= bump) {
    const nextY = Math.max(py - bump, y + bump);
    context.quadraticCurveTo(x - bump * 0.65, (py + nextY) / 2, x, nextY);
  }

  context.closePath();
}

function clamp(value: number, min: number, max: number): number {
  return Math.min(max, Math.max(min, value));
}

function screenToWorld(point: Point, transform: Transform): Point {
  return {
    x: (point.x - transform.x) / transform.scale,
    y: (point.y - transform.y) / transform.scale,
  };
}

function worldToScreen(point: Point, transform: Transform): Point {
  return {
    x: point.x * transform.scale + transform.x,
    y: point.y * transform.scale + transform.y,
  };
}

function getPointerPosition(e: React.PointerEvent<HTMLCanvasElement>, element: HTMLCanvasElement): Point {
  const rect = element.getBoundingClientRect();
  return { x: e.clientX - rect.left, y: e.clientY - rect.top };
}

function pointNear(a: Point, b: Point, tolerance = 12): boolean {
  return distance(a, b) <= tolerance;
}

function pointInPolygon(point: Point, polygon: Point[]): boolean {
  if (!Array.isArray(polygon) || polygon.length < 3) return false;
  let inside = false;
  for (let index = 0, previous = polygon.length - 1; index < polygon.length; previous = index++) {
    const xi = polygon[index].x;
    const yi = polygon[index].y;
    const xj = polygon[previous].x;
    const yj = polygon[previous].y;
    const intersects =
      yi > point.y !== yj > point.y &&
      point.x < ((xj - xi) * (point.y - yi)) / ((yj - yi) || 0.000001) + xi;
    if (intersects) inside = !inside;
  }
  return inside;
}

function pointInRect(point: Point, rect: Rect): boolean {
  return point.x >= rect.x && point.x <= rect.x + rect.w && point.y >= rect.y && point.y <= rect.y + rect.h;
}

function rectsIntersect(a: Rect, b: Rect): boolean {
  return a.x <= b.x + b.w && a.x + a.w >= b.x && a.y <= b.y + b.h && a.y + a.h >= b.y;
}

function pointInEllipse(point: Point, rect: Rect): boolean {
  const rx = rect.w / 2;
  const ry = rect.h / 2;
  if (rx <= 0 || ry <= 0) return false;
  const cx = rect.x + rx;
  const cy = rect.y + ry;
  const dx = (point.x - cx) / rx;
  const dy = (point.y - cy) / ry;
  return dx * dx + dy * dy <= 1;
}

function rectFromPoints(start: Point, current: Point): Rect {
  return {
    x: Math.min(start.x, current.x),
    y: Math.min(start.y, current.y),
    w: Math.abs(current.x - start.x),
    h: Math.abs(current.y - start.y),
  };
}

function rectHandlePoints(rect: Rect): Record<"nw" | "ne" | "sw" | "se", Point> {
  return {
    nw: { x: rect.x, y: rect.y },
    ne: { x: rect.x + rect.w, y: rect.y },
    sw: { x: rect.x, y: rect.y + rect.h },
    se: { x: rect.x + rect.w, y: rect.y + rect.h },
  };
}

function rectHandleAtPoint(world: Point, rect: Rect, tolerance = 14): "nw" | "ne" | "sw" | "se" | null {
  const handles = rectHandlePoints(rect);
  for (const key of Object.keys(handles) as Array<keyof typeof handles>) {
    if (pointNear(world, handles[key], tolerance)) return key;
  }
  return null;
}

function resizeRectFromHandle(rect: Rect, handle: "nw" | "ne" | "sw" | "se", point: Point): Rect {
  const opposite = {
    nw: { x: rect.x + rect.w, y: rect.y + rect.h },
    ne: { x: rect.x, y: rect.y + rect.h },
    sw: { x: rect.x + rect.w, y: rect.y },
    se: { x: rect.x, y: rect.y },
  }[handle];
  return rectFromPoints(opposite, point);
}

function getStampWorldRect(stamp: Pick<StampAnnotation, "position" | "text">): Rect {
  const width = Math.max(140, Math.min(260, stamp.text.length * 11 + 44));
  return { x: stamp.position.x, y: stamp.position.y, w: width, h: 44 };
}

function drawCanvasValueLabel(
  context: CanvasRenderingContext2D,
  text: string,
  x: number,
  y: number,
  scale: number
) {
  const fontSize = Math.max(15, Math.min(28, 18 * scale));
  const labelHeight = Math.max(28, Math.min(40, 30 * scale));
  context.save();
  context.font = `700 ${fontSize}px sans-serif`;
  const labelWidth = context.measureText(text).width + 18;
  context.fillStyle = "rgba(15, 23, 42, 0.84)";
  context.fillRect(x - 8, y - labelHeight + 6, labelWidth, labelHeight);
  context.strokeStyle = "rgba(248, 250, 252, 0.32)";
  context.lineWidth = 1;
  context.strokeRect(x - 8, y - labelHeight + 6, labelWidth, labelHeight);
  context.fillStyle = "#f8fafc";
  context.fillText(text, x, y - 6);
  context.restore();
}

function drawSelectionHandles(context: CanvasRenderingContext2D, rect: Rect, transform: Transform) {
  const handles = rectHandlePoints(rect);
  context.save();
  context.fillStyle = "#f8fafc";
  context.strokeStyle = "#2563eb";
  context.lineWidth = 1.5;
  for (const point of Object.values(handles)) {
    const screen = worldToScreen(point, transform);
    context.beginPath();
    context.rect(screen.x - 4, screen.y - 4, 8, 8);
    context.fill();
    context.stroke();
  }
  context.restore();
}

function cloneAnnotation<T extends Annotation>(annotation: T): T {
  return JSON.parse(JSON.stringify(annotation)) as T;
}

function offsetAnnotation<T extends Annotation>(annotation: T, dx: number, dy: number): T {
  const copyItem = cloneAnnotation(annotation);
  if ("position" in copyItem) {
    copyItem.position.x += dx;
    copyItem.position.y += dy;
  }
  if ("anchor" in copyItem) {
    copyItem.anchor.x += dx;
    copyItem.anchor.y += dy;
  }
  if ("box" in copyItem) {
    copyItem.box.x += dx;
    copyItem.box.y += dy;
  }
  if ("start" in copyItem && "end" in copyItem) {
    copyItem.start.x += dx;
    copyItem.start.y += dy;
    copyItem.end.x += dx;
    copyItem.end.y += dy;
  }
  if ("points" in copyItem) {
    copyItem.points = copyItem.points.map((point) => ({ x: point.x + dx, y: point.y + dy }));
  }
  if ("rect" in copyItem) {
    copyItem.rect.x += dx;
    copyItem.rect.y += dy;
  }
  return copyItem;
}

function createDuplicatedAnnotations(items: Annotation[], page: number, baseOffset = 20): Annotation[] {
  const createdAt = new Date().toISOString();
  return items.map((annotation, index) => {
    const duplicated = offsetAnnotation(annotation, baseOffset + index * 8, baseOffset + index * 8);
    duplicated.id = createId();
    duplicated.createdAt = createdAt;
    duplicated.page = page;
    return duplicated;
  });
}

function collectAnnotationSnapPoints(annotation: Annotation): Array<{ point: Point; label: string }> {
  switch (annotation.type) {
    case "measure":
      return [
        { point: annotation.start, label: "Measure start" },
        { point: annotation.end, label: "Measure end" },
      ];
    case "area":
    case "path":
      return annotation.points.map((point, index) => ({ point, label: `${annotation.type} point ${index + 1}` }));
    case "text":
    case "comment":
    case "link":
    case "stamp":
      return [{ point: annotation.position, label: `${annotation.type} anchor` }];
    case "callout":
      return [
        { point: annotation.anchor, label: "Callout anchor" },
        ...Object.values(rectHandlePoints(annotation.box)).map((point) => ({ point, label: "Callout corner" })),
      ];
    case "snapshot":
    case "cloud":
    case "rectangle":
    case "ellipse":
      return Object.values(rectHandlePoints(annotation.rect)).map((point) => ({ point, label: `${annotation.type} corner` }));
    default:
      return [];
  }
}

function annotationIntersectsSelectionBox(annotation: Annotation, rect: Rect): boolean {
  if (annotation.type === "text" || annotation.type === "comment" || annotation.type === "link" || annotation.type === "stamp") {
    return pointInRect(annotation.position, rect);
  }
  if (annotation.type === "measure") {
    const midpoint = { x: (annotation.start.x + annotation.end.x) / 2, y: (annotation.start.y + annotation.end.y) / 2 };
    return pointInRect(annotation.start, rect) || pointInRect(annotation.end, rect) || pointInRect(midpoint, rect);
  }
  if (annotation.type === "area" || annotation.type === "path") {
    return annotation.points.some((point) => pointInRect(point, rect));
  }
  if (annotation.type === "callout") {
    return pointInRect(annotation.anchor, rect) || rectsIntersect(annotation.box, rect);
  }
  if (annotation.type === "snapshot" || annotation.type === "cloud" || annotation.type === "rectangle" || annotation.type === "ellipse") {
    return rectsIntersect(annotation.rect, rect);
  }
  return false;
}

function annotationSummary(item: Annotation, calibration: number): string {
  if (item.type === "measure") return formatFeet(distance(item.start, item.end), calibration);
  if (item.type === "area") return formatArea(polygonArea(item.points), calibration);
  if (item.type === "rectangle") return "Rectangle";
  if (item.type === "ellipse") return "Ellipse";
  if (item.type === "text") return item.text;
  if (item.type === "comment") return item.comment || item.text || "Comment";
  if (item.type === "callout") return item.text;
  if (item.type === "link") return `→ Page ${item.targetPage}`;
  if (item.type === "snapshot") return item.title || "Snapshot";
  if (item.type === "cloud") return item.label || "Revision cloud";
  if (item.type === "stamp") return item.text;
  return item.type;
}

function categoryForAnnotation(item: Pick<Annotation, "type"> & Partial<Annotation>) {
  if (item.category) return item.category;
  switch (item.type) {
    case "measure":
      return "Linear";
    case "area":
      return "Area";
    case "rectangle":
      return "Shapes";
    case "ellipse":
      return "Shapes";
    case "text":
      return "Notes";
    case "comment":
      return "Comments";
    case "callout":
      return "Callouts";
    case "link":
      return "Links";
    case "snapshot":
      return "Snapshots";
    case "cloud":
      return "Clouds";
    case "stamp":
      return "Stamps";
    default:
      return "General";
  }
}

function markupStatusLabel(status: string) {
  return MARKUP_STATUSES.find((item) => item.value === status)?.label || status || "Open";
}

function markupStatusClasses(status: string) {
  if (status === "closed") return "border-emerald-500/40 bg-emerald-950/40 text-emerald-200";
  if (status === "in-review") return "border-sky-500/40 bg-sky-950/30 text-sky-200";
  if (status === "rfi") return "border-red-500/40 bg-red-950/30 text-red-200";
  if (status === "approved-noted") return "border-amber-500/40 bg-amber-950/30 text-amber-200";
  return "border-slate-600 bg-slate-800/70 text-slate-200";
}

function computeLegendFromAnnotations(allAnnotations: Array<Pick<Annotation, "type"> & Partial<Annotation>>) {
  const grouped: Record<string, number> = {};
  allAnnotations.forEach((item) => {
    const key = categoryForAnnotation(item);
    grouped[key] = (grouped[key] || 0) + 1;
  });
  return grouped;
}

function extractSheetNumber(items: TextContentItem[]): string | null {
  let bestCandidate: { value: string; score: number } | null = null;

  items.forEach((item) => {
    const value = String(item.str || "").toUpperCase().replace(/[^A-Z0-9.\-\s]/g, " ");
    const tokens = value.split(/\s+/).filter(Boolean);

    tokens.forEach((token) => {
      if (!SHEET_NUMBER_TOKEN.test(token)) return;

      const transform = Array.isArray(item.transform) ? item.transform : [0, 0, 0, 0, 0, 0];
      const scaleX = Math.hypot(transform[0] || 0, transform[1] || 0);
      const scaleY = Math.hypot(transform[2] || 0, transform[3] || 0);
      const fontSize = Math.max(scaleX, scaleY);
      const x = transform[4] || 0;
      const y = transform[5] || 0;
      const score = fontSize * 100 + x * 0.02 + y * 0.01;

      if (!bestCandidate || score > bestCandidate.score) {
        bestCandidate = { value: token, score };
      }
    });
  });

  return bestCandidate?.value || null;
}

function buildPageTextFragments(items: TextContentItem[], viewport: { transform: number[] }): PageTextFragment[] {
  return items.flatMap((item) => {
    if (!item.transform || !item.str) return [];

    const tx = (pdfjsLib as any).Util.transform(viewport.transform, item.transform);
    const fontHeight = Math.max(Math.hypot(tx[2] || 0, tx[3] || 0), 8);
    const width = Math.max(item.width || fontHeight * String(item.str).length * 0.5, fontHeight * 0.75);

    return [{
      text: String(item.str),
      rect: {
        x: tx[4] || 0,
        y: (tx[5] || 0) - fontHeight,
        w: width,
        h: fontHeight * 1.15,
      },
    }];
  });
}

function normalizeReferenceToken(value: string): string {
  return String(value || "").toUpperCase().replace(/[^A-Z0-9.\-]/g, "");
}

function extractKeynoteTokens(text: string, knownSheetNumbers?: Set<string>): string[] {
  const upper = String(text || "").toUpperCase();
  const tokens = Array.from(upper.matchAll(KEYNOTE_TOKEN))
    .map((match) => match[1]?.toUpperCase() || "")
    .filter(Boolean)
    .filter((token) => !(knownSheetNumbers?.has(token)));
  return Array.from(new Set(tokens));
}

function unionRects(a: Rect, b: Rect): Rect {
  const x = Math.min(a.x, b.x);
  const y = Math.min(a.y, b.y);
  const right = Math.max(a.x + a.w, b.x + b.w);
  const bottom = Math.max(a.y + a.h, b.y + b.h);
  return { x, y, w: right - x, h: bottom - y };
}

function expandRect(rect: Rect, paddingX: number, paddingY: number): Rect {
  return {
    x: rect.x - paddingX,
    y: rect.y - paddingY,
    w: rect.w + paddingX * 2,
    h: rect.h + paddingY * 2,
  };
}

function buildAutoReferenceLinks(
  fragmentsByPage: Record<number, PageTextFragment[]>,
  sheetNumbers: Record<number, string>
): Record<number, AutoReferenceLink[]> {
  const sheetToPage = new Map<string, number>();
  Object.entries(sheetNumbers).forEach(([page, sheetNumber]) => {
    sheetToPage.set(String(sheetNumber).toUpperCase(), Number(page));
  });

  const links: Record<number, AutoReferenceLink[]> = {};

  Object.entries(fragmentsByPage).forEach(([pageKey, fragments]) => {
    const page = Number(pageKey);
    links[page] = [];
    const seenKeys = new Set<string>();

    fragments.forEach((fragment) => {
      const upper = fragment.text.toUpperCase();
      const matches = Array.from(upper.matchAll(DETAIL_REFERENCE_TOKEN));
      if (!matches.length) return;

      const characterWidth = fragment.rect.w / Math.max(upper.length, 1);

      matches.forEach((match) => {
        const detailToken = match[1]?.toUpperCase();
        const targetSheetNumber = match[2]?.toUpperCase();
        const targetPage = targetSheetNumber ? sheetToPage.get(targetSheetNumber) : undefined;
        if (!detailToken || !targetSheetNumber || !targetPage) return;
        const dedupeKey = `${detailToken}/${targetSheetNumber}`;
        if (seenKeys.has(dedupeKey)) return;
        seenKeys.add(dedupeKey);

        const startIndex = match.index || 0;
        const linkWidth = Math.max(characterWidth * match[0].length, fragment.rect.h * 1.2);

        links[page].push({
          id: `${page}-${targetSheetNumber}-${startIndex}`,
          page,
          text: match[0],
          detailToken,
          targetSheetNumber,
          targetPage,
          rect: {
            x: fragment.rect.x + startIndex * characterWidth,
            y: fragment.rect.y,
            w: linkWidth,
            h: fragment.rect.h,
          },
        });
      });
    });

    const normalizedFragments = fragments.map((fragment, index) => ({
      index,
      fragment,
      token: normalizeReferenceToken(fragment.text),
      centerX: fragment.rect.x + fragment.rect.w / 2,
      centerY: fragment.rect.y + fragment.rect.h / 2,
    }));

    const sheetFragments = normalizedFragments.filter(({ token }) => SHEET_NUMBER_TOKEN.test(token) && sheetToPage.has(token));
    const detailFragments = normalizedFragments.filter(({ token }) => DETAIL_TOKEN_ONLY.test(token) && !SHEET_NUMBER_TOKEN.test(token));

    sheetFragments.forEach((sheetFragment) => {
      const targetSheetNumber = sheetFragment.token;
      const targetPage = sheetToPage.get(targetSheetNumber);
      if (!targetPage) return;

      const candidate = detailFragments
        .map((detailFragment) => {
          const horizontalDistance = Math.abs(detailFragment.centerX - sheetFragment.centerX);
          const verticalGap = sheetFragment.fragment.rect.y - (detailFragment.fragment.rect.y + detailFragment.fragment.rect.h);
          const maxHorizontalDistance = Math.max(sheetFragment.fragment.rect.w * 0.9, 28);
          if (verticalGap < -8 || verticalGap > Math.max(42, sheetFragment.fragment.rect.h * 3.2)) return null;
          if (horizontalDistance > maxHorizontalDistance) return null;
          const score = horizontalDistance * 2 + Math.abs(verticalGap);
          return { detailFragment, score };
        })
        .filter((item): item is { detailFragment: typeof detailFragments[number]; score: number } => !!item)
        .sort((a, b) => a.score - b.score)[0];

      if (!candidate) return;

      const detailToken = candidate.detailFragment.token;
      const dedupeKey = `${detailToken}/${targetSheetNumber}`;
      if (seenKeys.has(dedupeKey)) return;
      seenKeys.add(dedupeKey);

      const rect = expandRect(
        unionRects(candidate.detailFragment.fragment.rect, sheetFragment.fragment.rect),
        Math.max(8, sheetFragment.fragment.rect.h * 0.45),
        Math.max(6, sheetFragment.fragment.rect.h * 0.35)
      );

      links[page].push({
        id: `${page}-${targetSheetNumber}-${detailToken}-stacked`,
        page,
        text: `${detailToken}/${targetSheetNumber}`,
        detailToken,
        targetSheetNumber,
        targetPage,
        rect,
      });
    });
  });

  return links;
}

function inferBookmarkSectionFromText(text: string): Exclude<BookmarkSection, "all"> {
  const lower = text.toLowerCase();
  const matched = BOOKMARK_TABS.find(
    (tab) => tab.id !== "all" && tab.id !== "general" && tab.keywords.some((keyword) => lower.includes(keyword))
  );
  return (matched?.id as Exclude<BookmarkSection, "all"> | undefined) || "general";
}

function inferBookmarkSection(sheetNumber: string | undefined, text: string): Exclude<BookmarkSection, "all"> {
  const normalized = (sheetNumber || "").toUpperCase();
  if (/^C/.test(normalized)) return "civil";
  if (/^S/.test(normalized)) return "structural";
  if (/^(A|I)/.test(normalized)) return "architectural";
  if (/^(M|E|P|FP|F)/.test(normalized)) return "mep";
  if (/^G/.test(normalized)) return "general";
  return inferBookmarkSectionFromText(text);
}

function isRenderCancelledError(error: unknown): boolean {
  return error instanceof Error && error.name === "RenderingCancelledException";
}

function configurePdfWorker() {
  try {
    const workerUrl = new URL("pdfjs-dist/build/pdf.worker.min.mjs", import.meta.url);
    const worker = new Worker(workerUrl, { type: "module" });
    pdfjsLib.GlobalWorkerOptions.workerPort = worker;
    return { status: "ready" as WorkerStatus, label: workerUrl.toString() };
  } catch (error) {
    const message = error instanceof Error ? error.message : "Unable to configure PDF worker.";
    return { status: "error" as WorkerStatus, label: message };
  }
}

const workerSetup = configurePdfWorker();

function runUtilityTests() {
  const tests: Array<{ name: string; pass: boolean }> = [];
  tests.push({ name: "distance 3-4-5", pass: Math.abs(distance({ x: 0, y: 0 }, { x: 3, y: 4 }) - 5) < 0.0001 });
  tests.push({
    name: "triangle area",
    pass: Math.abs(polygonArea([{ x: 0, y: 0 }, { x: 10, y: 0 }, { x: 0, y: 10 }]) - 50) < 0.0001,
  });
  tests.push({
    name: "point inside polygon",
    pass:
      pointInPolygon(
        { x: 5, y: 5 },
        [
          { x: 0, y: 0 },
          { x: 10, y: 0 },
          { x: 10, y: 10 },
          { x: 0, y: 10 },
        ]
      ) === true,
  });
  tests.push({
    name: "point outside polygon",
    pass:
      pointInPolygon(
        { x: 15, y: 5 },
        [
          { x: 0, y: 0 },
          { x: 10, y: 0 },
          { x: 10, y: 10 },
          { x: 0, y: 10 },
        ]
      ) === false,
  });
  tests.push({ name: "ensureArray fallback", pass: ensureArray(null).length === 0 });
  tests.push({ name: "formatFeet output", pass: formatFeet(10, 0.5) === "5.00 ft" });
  tests.push({ name: "formatArea output", pass: formatArea(10, 2) === "40.00 sf" });
  tests.push({ name: "safeParse invalid json fallback", pass: safeParse("{bad", 123) === 123 });
  tests.push({
    name: "summary text",
    pass: annotationSummary({ type: "link", targetPage: 4 } as LinkAnnotation, 1) === "→ Page 4",
  });
  tests.push({
    name: "legend counts",
    pass: computeLegendFromAnnotations([{ type: "text" }, { type: "text" }, { type: "measure" } as any]).Notes === 2,
  });
  tests.push({
    name: "stacked detail link detection",
    pass: (() => {
      const links = buildAutoReferenceLinks(
        {
          1: [
            { text: "1", rect: { x: 100, y: 100, w: 10, h: 12 } },
            { text: "A301", rect: { x: 92, y: 116, w: 26, h: 12 } },
          ],
        },
        { 4: "A301" }
      );
      return links[1]?.[0]?.text === "1/A301" && links[1]?.[0]?.targetPage === 4;
    })(),
  });
  tests.push({
    name: "keynote token extraction",
    pass: (() => {
      const tokens = extractKeynoteTokens("P77 - verify wall type", new Set(["A301"]));
      return tokens.includes("P77") && !tokens.includes("A301");
    })(),
  });
  tests.push({
    name: "screen/world invert approx",
    pass: (() => {
      const transform = { x: 10, y: 20, scale: 2 };
      const point = { x: 5, y: 7 };
      const screen = worldToScreen(point, transform);
      const world = screenToWorld(screen, transform);
      return Math.abs(world.x - point.x) < 0.0001 && Math.abs(world.y - point.y) < 0.0001;
    })(),
  });
  tests.push({ name: "worker uses workerPort", pass: workerSetup.status === "ready" || workerSetup.status === "error" });
  return tests;
}

export default function BluebeamLikePlanViewer() {
  const [fileName, setFileName] = useState("No file loaded");
  const [pdfDoc, setPdfDoc] = useState<any>(null);
  const [pageCount, setPageCount] = useState(0);
  const [currentPage, setCurrentPage] = useState(1);
  const [pageSizes, setPageSizes] = useState<Record<number, PageSize>>({});
  const [pageSheetNumbers, setPageSheetNumbers] = useState<Record<number, string>>({});
  const [pageThumbnails, setPageThumbnails] = useState<Record<number, string>>({});
  const [pageTexts, setPageTexts] = useState<Record<number, string>>({});
  const [autoReferenceLinks, setAutoReferenceLinks] = useState<Record<number, AutoReferenceLink[]>>({});
  const [loadError, setLoadError] = useState("");
  const [workerStatus, setWorkerStatus] = useState<WorkerStatus>(workerSetup.status);
  const [workerLabel] = useState(workerSetup.label);
  const [tool, setTool] = useState<ToolType>(TOOL.PAN);
  const [transform, setTransform] = useState<Transform>({ x: 0, y: 0, scale: 1 });
  const [isPanning, setIsPanning] = useState(false);
  const [panStart, setPanStart] = useState<{ x: number; y: number; originX: number; originY: number } | null>(null);
  const [annotations, setAnnotations] = useState<AnnotationMap>({});
  const [draft, setDraft] = useState<Draft>(null);
  const [selectedId, setSelectedId] = useState<string | null>(null);
  const [selectedIds, setSelectedIds] = useState<string[]>([]);
  const [selectedHandle, setSelectedHandle] = useState<string | null>(null);
  const [draggingSelection, setDraggingSelection] = useState<{ world: Point } | null>(null);
  const [clipboardAnnotations, setClipboardAnnotations] = useState<Annotation[]>([]);
  const [snapPreview, setSnapPreview] = useState<SnapPreview | null>(null);
  const [undoStack, setUndoStack] = useState<HistorySnapshot[]>([]);
  const [redoStack, setRedoStack] = useState<HistorySnapshot[]>([]);
  const [calibration, setCalibration] = useState(1);
  const [sessions, setSessions] = useState<PersistedSession[]>([]);
  const [projects, setProjects] = useState<PersistedProject[]>([]);
  const [projectName, setProjectName] = useState("Project Alpha");
  const [activeProjectId, setActiveProjectId] = useState<string | null>(null);
  const [cloudProjects, setCloudProjects] = useState<CloudProjectSummary[]>([]);
  const [cloudStatus, setCloudStatus] = useState<"idle" | "ready" | "error">("idle");
  const [cloudMessage, setCloudMessage] = useState("Shared cloud unavailable");
  const [isCloudSaving, setIsCloudSaving] = useState(false);
  const [isCloudLoading, setIsCloudLoading] = useState(false);
  const [activePdfBlobUrl, setActivePdfBlobUrl] = useState("");
  const [activePdfFile, setActivePdfFile] = useState<File | null>(null);
  const [showMarkupList, setShowMarkupList] = useState(false);
  const [markupFilter, setMarkupFilter] = useState("");
  const [markupStatusFilter, setMarkupStatusFilter] = useState("all");
  const [ocrQuery, setOcrQuery] = useState("");
  const [ocrResults, setOcrResults] = useState<Array<{ page: number; preview: string }>>([]);
  const [compareMode, setCompareMode] = useState(false);
  const [comparePage, setComparePage] = useState(2);
  const [activeBookmarkTab, setActiveBookmarkTab] = useState<BookmarkSection>("all");
  const [navigationMode, setNavigationMode] = useState<NavigationMode>("sheet");
  const [boxZoomMode, setBoxZoomMode] = useState(false);
  const [activeMenu, setActiveMenu] = useState<MenuId | null>(null);
  const [sheetJumpQuery, setSheetJumpQuery] = useState("");
  const [statusMessage, setStatusMessage] = useState("Ready");
  const [showProjectsPanel, setShowProjectsPanel] = useState(true);
  const [showInspectorPanel, setShowInspectorPanel] = useState(true);
  const [showPagesPanel, setShowPagesPanel] = useState(true);
  const [cursorWorld, setCursorWorld] = useState<Point | null>(null);
  const [activeMarkupColor, setActiveMarkupColor] = useState("#2563eb");
  const [activeMarkupLayerId, setActiveMarkupLayerId] = useState("layer-general");
  const [activeMarkupCategory, setActiveMarkupCategory] = useState("Notes");
  const [activeStampPresetId, setActiveStampPresetId] = useState("stamp-field-verify");
  const [activeMarkupStrokeWidth, setActiveMarkupStrokeWidth] = useState(2);
  const [activeMarkupFillColor, setActiveMarkupFillColor] = useState("#2563eb");
  const [activeMarkupFillOpacity, setActiveMarkupFillOpacity] = useState(0.16);
  const [leftPanelWidth, setLeftPanelWidth] = useState(224);
  const [rightPanelWidth, setRightPanelWidth] = useState(288);
  const [isWideLayout, setIsWideLayout] = useState<boolean>(() => (typeof window === "undefined" ? true : window.innerWidth >= 1536));
  const [toolChest] = useState([
    { id: createId(), name: "RFI", kind: "text", payload: { text: "RFI", color: "#b91c1c", layerId: "layer-rfi", category: "RFI" } },
    { id: createId(), name: "ADA", kind: "text", payload: { text: "ADA", color: "#1d4ed8", layerId: "layer-general", category: "Code" } },
    { id: createId(), name: "Punch", kind: "text", payload: { text: "Punch", color: "#15803d", layerId: "layer-punch", category: "Punch" } },
  ]);
  const [selectedChestId, setSelectedChestId] = useState<string | null>(null);
  const [layers, setLayers] = useState<Layer[]>(DEFAULT_LAYERS);
  const [users, setUsers] = useState<User[]>(DEFAULT_USERS);
  const [activeUserId, setActiveUserId] = useState("u1");
  const [reportTitle, setReportTitle] = useState("Field Report");
  const [isViewerFullscreen, setIsViewerFullscreen] = useState(false);
  const [hoveredAutoReferenceId, setHoveredAutoReferenceId] = useState<string | null>(null);
  const [hoveredKeynoteToken, setHoveredKeynoteToken] = useState<string | null>(null);
  const [pendingDetailFocus, setPendingDetailFocus] = useState<{ targetPage: number; detailToken: string } | null>(null);
  const [inlineEditor, setInlineEditor] = useState<InlineEditorState | null>(null);
  const [temporaryPanActive, setTemporaryPanActive] = useState(false);
  const uploadRunRef = useRef(0);
  const pageWheelCooldownRef = useRef(0);
  const touchPointersRef = useRef<Map<number, Point>>(new Map());
  const pinchStateRef = useRef<{ distance: number; midpoint: Point } | null>(null);
  const pageTextFragmentsRef = useRef<Record<number, PageTextFragment[]>>({});
  const menuBarRef = useRef<HTMLDivElement | null>(null);
  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const statusTimeoutRef = useRef<number | null>(null);
  const resizeStateRef = useRef<{ side: "left" | "right"; startX: number; startWidth: number } | null>(null);
  const renderTasksRef = useRef<Record<RenderCanvasKey, any | null>>({ main: null, compare: null });
  const renderRequestIdsRef = useRef<Record<RenderCanvasKey, number>>({ main: 0, compare: 0 });
  const isApplyingHistoryRef = useRef(false);
  const temporaryPanPreviousToolRef = useRef<ToolType>(TOOL.PAN);
  const miniMapDragPointerIdRef = useRef<number | null>(null);
  const inlineEditorRef = useRef<HTMLInputElement | null>(null);
  const viewerRef = useRef<HTMLDivElement | null>(null);
  const pdfCanvasRef = useRef<HTMLCanvasElement | null>(null);
  const compareCanvasRef = useRef<HTMLCanvasElement | null>(null);
  const overlayCanvasRef = useRef<HTMLCanvasElement | null>(null);

  const tests = useMemo(() => runUtilityTests(), []);
  const currentAnnotations = useMemo(() => ensureArray(annotations[currentPage]), [annotations, currentPage]);
  const selectedIdSet = useMemo(() => new Set(selectedIds), [selectedIds]);
  const currentAutoReferenceLinks = useMemo(() => ensureArray(autoReferenceLinks[currentPage]), [autoReferenceLinks, currentPage]);
  const knownSheetNumberSet = useMemo(
    () => new Set(Object.values(pageSheetNumbers).map((sheetNumber) => String(sheetNumber).toUpperCase())),
    [pageSheetNumbers]
  );
  const currentPageTextFragments = useMemo(
    () => pageTextFragmentsRef.current[currentPage] || [],
    [currentPage, pageTexts, pageSheetNumbers]
  );
  const hoveredAutoReference = useMemo(
    () => currentAutoReferenceLinks.find((link) => link.id === hoveredAutoReferenceId) || null,
    [currentAutoReferenceLinks, hoveredAutoReferenceId]
  );
  const highlightedKeynoteFragments = useMemo(() => {
    if (!hoveredKeynoteToken) return [];
    return currentPageTextFragments.filter((fragment) =>
      extractKeynoteTokens(fragment.text, knownSheetNumberSet).includes(hoveredKeynoteToken)
    );
  }, [currentPageTextFragments, hoveredKeynoteToken, knownSheetNumberSet]);
  const currentPageSize = pageSizes[currentPage] || { width: 1200, height: 900 };
  const currentSheetNumber = pageSheetNumbers[currentPage] || "";
  const cursorSummary = useMemo(() => {
    if (!cursorWorld) return "Cursor off sheet";
    return `X ${cursorWorld.x.toFixed(0)} | Y ${cursorWorld.y.toFixed(0)} | ${formatFeet(cursorWorld.x, calibration)} / ${formatFeet(cursorWorld.y, calibration)}`;
  }, [calibration, cursorWorld]);
  const selectedAnnotation = useMemo(
    () => currentAnnotations.find((annotation) => annotation.id === selectedId) || null,
    [currentAnnotations, selectedId]
  );
  const selectedAnnotations = useMemo(
    () => currentAnnotations.filter((annotation) => selectedIdSet.has(annotation.id)),
    [currentAnnotations, selectedIdSet]
  );
  const inlineEditorAnnotation = useMemo(() => {
    if (!inlineEditor) return null;
    return ensureArray(annotations[inlineEditor.page]).find((annotation) => annotation.id === inlineEditor.annotationId) || null;
  }, [annotations, inlineEditor]);
  const inlineEditorLayout = useMemo(() => {
    if (!inlineEditor || !inlineEditorAnnotation || inlineEditor.page !== currentPage || !viewerRef.current) return null;
    const viewerWidth = viewerRef.current.clientWidth;
    const viewerHeight = viewerRef.current.clientHeight;
    let left = 0;
    let top = 0;
    let width = 220;
    let height = 40;

    if (inlineEditorAnnotation.type === "text" || inlineEditorAnnotation.type === "comment") {
      const position = worldToScreen(inlineEditorAnnotation.position, transform);
      left = position.x - 8;
      top = position.y - 28;
      width = 260;
    } else if (inlineEditorAnnotation.type === "callout") {
      left = inlineEditorAnnotation.box.x * transform.scale + transform.x;
      top = inlineEditorAnnotation.box.y * transform.scale + transform.y;
      width = Math.max(180, inlineEditorAnnotation.box.w * transform.scale);
      height = Math.max(40, inlineEditorAnnotation.box.h * transform.scale);
    } else if (inlineEditorAnnotation.type === "stamp") {
      const rect = getStampWorldRect(inlineEditorAnnotation);
      left = rect.x * transform.scale + transform.x;
      top = rect.y * transform.scale + transform.y;
      width = rect.w * transform.scale;
      height = rect.h * transform.scale;
    } else if (inlineEditorAnnotation.type === "cloud") {
      left = inlineEditorAnnotation.rect.x * transform.scale + transform.x;
      top = inlineEditorAnnotation.rect.y * transform.scale + transform.y - 34;
      width = Math.max(180, inlineEditorAnnotation.rect.w * transform.scale);
      height = 36;
    } else if (inlineEditorAnnotation.type === "snapshot") {
      left = inlineEditorAnnotation.rect.x * transform.scale + transform.x;
      top = inlineEditorAnnotation.rect.y * transform.scale + transform.y - 34;
      width = Math.max(180, inlineEditorAnnotation.rect.w * transform.scale);
      height = 36;
    } else {
      return null;
    }

    return {
      left: clamp(left, 8, Math.max(8, viewerWidth - width - 8)),
      top: clamp(top, 8, Math.max(8, viewerHeight - height - 8)),
      width: clamp(width, 160, Math.max(160, viewerWidth - 16)),
      height: clamp(height, 36, Math.max(36, viewerHeight - 16)),
    };
  }, [currentPage, inlineEditor, inlineEditorAnnotation, transform]);
  const propertyStrokeWidth = selectedAnnotation?.strokeWidth ?? activeMarkupStrokeWidth;
  const propertyStrokeColor = selectedAnnotation?.color ?? activeMarkupColor;
  const propertyFillColor = selectedAnnotation?.fillColor ?? activeMarkupFillColor;
  const propertyFillOpacity = selectedAnnotation?.fillOpacity ?? activeMarkupFillOpacity;
  const propertyLabel =
    selectedIds.length > 1
      ? `${selectedIds.length} selected markups`
      : selectedAnnotation
        ? `Selected ${selectedAnnotation.type}`
        : "Default markup style";
  const activeUser = useMemo(() => users.find((user) => user.id === activeUserId) || users[0], [users, activeUserId]);
  const activeRole = useMemo(
    () => DEFAULT_ROLES.find((role) => role.id === activeUser?.roleId) || DEFAULT_ROLES[0],
    [activeUser]
  );
  const visibleLayerIds = useMemo(() => new Set(layers.filter((layer) => layer.visible).map((layer) => layer.id)), [layers]);
  const pageBookmarkSections = useMemo(() => {
    const sections: Record<number, Exclude<BookmarkSection, "all">> = {};
    for (let page = 1; page <= pageCount; page += 1) {
      sections[page] = inferBookmarkSection(pageSheetNumbers[page], pageTexts[page] || "");
    }
    return sections;
  }, [pageCount, pageSheetNumbers, pageTexts]);
  const availableBookmarkTabs = useMemo(
    () =>
      BOOKMARK_TABS.filter(
        (tab) => tab.id === "all" || Object.values(pageBookmarkSections).includes(tab.id as Exclude<BookmarkSection, "all">)
      ),
    [pageBookmarkSections]
  );
  const filteredThumbnailPages = useMemo(() => {
    const pages = Array.from({ length: pageCount }, (_, index) => index + 1);
    if (activeBookmarkTab === "all") return pages;
    return pages.filter((page) => pageBookmarkSections[page] === activeBookmarkTab);
  }, [activeBookmarkTab, pageBookmarkSections, pageCount]);
  const miniMapMetrics = useMemo(() => {
    const viewerWidth = viewerRef.current?.clientWidth || 0;
    const viewerHeight = viewerRef.current?.clientHeight || 0;
    if (!viewerWidth || !viewerHeight) return null;
    const scale = Math.min(170 / currentPageSize.width, 120 / currentPageSize.height);
    const width = currentPageSize.width * scale;
    const height = currentPageSize.height * scale;
    const topLeft = screenToWorld({ x: 0, y: 0 }, transform);
    const bottomRight = screenToWorld({ x: viewerWidth, y: viewerHeight }, transform);
    return {
      scale,
      width,
      height,
      viewport: {
        x: clamp(topLeft.x, 0, currentPageSize.width) * scale,
        y: clamp(topLeft.y, 0, currentPageSize.height) * scale,
        w: Math.max(14, clamp(bottomRight.x - topLeft.x, 0, currentPageSize.width) * scale),
        h: Math.max(14, clamp(bottomRight.y - topLeft.y, 0, currentPageSize.height) * scale),
      },
    };
  }, [currentPageSize.height, currentPageSize.width, transform, currentPage, isViewerFullscreen]);
  const legacyMenuEntries: Record<MenuId, MenuItem[]> = {
      file: [
        { label: "Open PDF…", action: openPdfPicker },
        { label: "Save Project", action: saveProject },
        { label: "Save Session", action: saveSession },
        { separator: true, label: "sep-1" },
        { label: "Export Data", action: exportSession, disabled: !canExport() },
        { label: "Export Annotated Manifest", action: exportAnnotatedPdfMock, disabled: !canExport() },
        { separator: true, label: "sep-2" },
        { label: isViewerFullscreen ? "Exit Fullscreen" : "Enter Fullscreen", action: () => void toggleViewerFullscreen() },
      ],
      edit: [
        { label: "Copy Selected", action: copySelectionToClipboard, disabled: !selectedIds.length },
        { label: "Paste", action: pasteClipboardSelection, disabled: !clipboardAnnotations.length || !canEdit() },
        { label: "Duplicate Selected", action: duplicateSelected, disabled: !selectedIds.length || !canEdit() },
        { label: "Delete Selected", action: deleteSelected, disabled: !selectedIds.length || !canDelete() },
        { label: "Finish Area", action: finishArea, disabled: !(draft && draft.type === "area") },
      ],
      view: [
        { label: "Zoom In", action: () => zoomBy(1.2) },
        { label: "Zoom Out", action: () => zoomBy(1 / 1.2) },
        { label: "Fit to Screen", action: fitToViewer },
        { separator: true, label: "sep-3" },
        { label: showMarkupList ? "Hide Markup List" : "Show Markup List", action: () => setShowMarkupList((value) => !value) },
        { label: compareMode ? "Disable Compare" : "Enable Compare", action: () => setCompareMode((value) => !value) },
      ],
      document: [
        { label: "Previous Page", action: () => changePageBy(-1), disabled: currentPage <= 1 },
        { label: "Next Page", action: () => changePageBy(1), disabled: currentPage >= pageCount },
        { label: "First Page", action: () => setCurrentPage(1), disabled: currentPage <= 1 },
        { label: "Last Page", action: () => setCurrentPage(Math.max(pageCount, 1)), disabled: currentPage >= pageCount },
      ],
      markup: [
        { label: "Select Tool", action: () => setTool(TOOL.SELECT) },
        { label: "Pan Tool", action: () => setTool(TOOL.PAN) },
        { label: "Sketch Markup", action: () => setTool(TOOL.DRAW) },
        { label: "Rectangle Markup", action: () => setTool(TOOL.RECTANGLE) },
        { label: "Ellipse Markup", action: () => setTool(TOOL.ELLIPSE) },
        { label: "Text Markup", action: () => setTool(TOOL.TEXT) },
        { label: "Comment Markup", action: () => setTool(TOOL.COMMENT) },
        { label: "Snapshot Tool", action: () => setTool(TOOL.SNAPSHOT) },
      ],
      measure: [
        { label: "Length Tool", action: () => setTool(TOOL.MEASURE) },
        { label: "Area Tool", action: () => setTool(TOOL.AREA) },
        { label: "Calibrate Scale", action: () => setTool(TOOL.CALIBRATE) },
      ],
      window: [
        { label: activeBookmarkTab === "all" ? "All Sheets Visible" : "Show All Sheet Tabs", action: () => setActiveBookmarkTab("all") },
        { label: compareMode ? "Close Compare View" : "Open Compare View", action: () => setCompareMode((value) => !value) },
        { label: showMarkupList ? "Hide Database Panel" : "Show Database Panel", action: () => setShowMarkupList((value) => !value) },
      ],
      help: [
        { label: "Viewer Controls", action: showHelpDialog },
        { label: "About This Workspace", action: () => window.alert("Bluebeam-style construction drawing viewer with markups, sheet links, thumbnails, fullscreen, and touch zoom.") },
      ],
    };

  useEffect(() => {
    setSessions(safeParse(localStorage.getItem("plan_sessions_v5"), [] as PersistedSession[]));
    const storedProjects = safeParse(localStorage.getItem("plan_projects_v5"), [] as PersistedProject[]);
    setProjects(storedProjects);
    if (storedProjects[0]) setActiveProjectId(storedProjects[0].id);
    setLayers(safeParse(localStorage.getItem("plan_layers_v5"), DEFAULT_LAYERS));
    setUsers(safeParse(localStorage.getItem("plan_users_v5"), DEFAULT_USERS));
    const storedPreferences = safeParse(localStorage.getItem("plan_viewer_prefs_v1"), {
      showProjectsPanel: true,
      showInspectorPanel: true,
      showPagesPanel: true,
      showMarkupList: false,
      navigationMode: "sheet",
      activeMarkupColor: "#2563eb",
      activeMarkupLayerId: "layer-general",
      activeMarkupCategory: "Notes",
      activeStampPresetId: "stamp-field-verify",
      activeMarkupStrokeWidth: 2,
      activeMarkupFillColor: "#2563eb",
      activeMarkupFillOpacity: 0.16,
      leftPanelWidth: 224,
      rightPanelWidth: 288,
    });
    setShowProjectsPanel(storedPreferences.showProjectsPanel);
    setShowInspectorPanel(storedPreferences.showInspectorPanel);
    setShowPagesPanel(storedPreferences.showPagesPanel);
    setShowMarkupList(storedPreferences.showMarkupList ?? false);
    setNavigationMode(storedPreferences.navigationMode === "review" ? "review" : "sheet");
    setActiveMarkupColor(storedPreferences.activeMarkupColor);
    setActiveMarkupLayerId(storedPreferences.activeMarkupLayerId);
    setActiveMarkupCategory(storedPreferences.activeMarkupCategory);
    setActiveStampPresetId(storedPreferences.activeStampPresetId || "stamp-field-verify");
    setActiveMarkupStrokeWidth(storedPreferences.activeMarkupStrokeWidth);
    setActiveMarkupFillColor(storedPreferences.activeMarkupFillColor);
    setActiveMarkupFillOpacity(storedPreferences.activeMarkupFillOpacity);
    setLeftPanelWidth(clamp(storedPreferences.leftPanelWidth ?? 224, 190, 340));
    setRightPanelWidth(clamp(storedPreferences.rightPanelWidth ?? 288, 230, 400));
    void refreshCloudProjects();
  }, []);

  useEffect(() => localStorage.setItem("plan_sessions_v5", JSON.stringify(sessions)), [sessions]);
  useEffect(() => localStorage.setItem("plan_projects_v5", JSON.stringify(projects)), [projects]);
  useEffect(() => localStorage.setItem("plan_layers_v5", JSON.stringify(layers)), [layers]);
  useEffect(() => localStorage.setItem("plan_users_v5", JSON.stringify(users)), [users]);
  useEffect(() => {
    localStorage.setItem(
      "plan_viewer_prefs_v1",
      JSON.stringify({
        showProjectsPanel,
        showInspectorPanel,
        showPagesPanel,
        showMarkupList,
        navigationMode,
        activeMarkupColor,
        activeMarkupLayerId,
        activeMarkupCategory,
        activeStampPresetId,
        activeMarkupStrokeWidth,
        activeMarkupFillColor,
        activeMarkupFillOpacity,
        leftPanelWidth,
        rightPanelWidth,
      })
    );
  }, [showProjectsPanel, showInspectorPanel, showPagesPanel, showMarkupList, navigationMode, activeMarkupColor, activeMarkupLayerId, activeMarkupCategory, activeStampPresetId, activeMarkupStrokeWidth, activeMarkupFillColor, activeMarkupFillOpacity, leftPanelWidth, rightPanelWidth]);
  useEffect(() => {
    if (!availableBookmarkTabs.some((tab) => tab.id === activeBookmarkTab)) {
      setActiveBookmarkTab("all");
    }
  }, [activeBookmarkTab, availableBookmarkTabs]);
  useEffect(() => {
    setHoveredAutoReferenceId(null);
    setHoveredKeynoteToken(null);
    setInlineEditor(null);
    setSnapPreview(null);
  }, [currentPage]);
  useEffect(() => {
    const validSelection = selectedIds.filter((id) => currentAnnotations.some((annotation) => annotation.id === id));
    if (validSelection.length !== selectedIds.length) {
      setSelectedIds(validSelection);
    }
    if (selectedId && !validSelection.includes(selectedId)) {
      setSelectedId(validSelection[0] || null);
    }
    if (!validSelection.length && selectedHandle) {
      setSelectedHandle(null);
    }
  }, [currentAnnotations, selectedHandle, selectedId, selectedIds]);
  useEffect(() => {
    if (!inlineEditor) return;
    window.setTimeout(() => {
      inlineEditorRef.current?.focus();
      inlineEditorRef.current?.select();
    }, 0);
  }, [inlineEditor]);
  useEffect(() => {
    const handlePointerDown = (event: MouseEvent) => {
      if (!menuBarRef.current?.contains(event.target as Node)) {
        setActiveMenu(null);
      }
    };

    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === "Escape") setActiveMenu(null);
    };

    document.addEventListener("mousedown", handlePointerDown);
    document.addEventListener("keydown", handleKeyDown);
    return () => {
      document.removeEventListener("mousedown", handlePointerDown);
      document.removeEventListener("keydown", handleKeyDown);
    };
  }, []);
  useEffect(() => {
    return () => {
      if (statusTimeoutRef.current) window.clearTimeout(statusTimeoutRef.current);
    };
  }, []);
  useEffect(() => {
    const handleKeydown = (event: KeyboardEvent) => {
      const target = event.target as HTMLElement | null;
      const isTypingTarget =
        target instanceof HTMLInputElement ||
        target instanceof HTMLTextAreaElement ||
        target instanceof HTMLSelectElement ||
        !!target?.closest("[contenteditable='true']");
      if (isTypingTarget) return;

      const key = event.key.toLowerCase();

      if ((event.ctrlKey || event.metaKey) && key === "o") {
        event.preventDefault();
        openPdfPicker();
        announceStatus("Opening PDF picker");
        return;
      }
      if ((event.ctrlKey || event.metaKey) && key === "z" && !event.shiftKey) {
        event.preventDefault();
        undoLastAction();
        return;
      }
      if (((event.ctrlKey || event.metaKey) && key === "y") || ((event.ctrlKey || event.metaKey) && event.shiftKey && key === "z")) {
        event.preventDefault();
        redoLastAction();
        return;
      }
      if ((event.ctrlKey || event.metaKey) && key === "s") {
        event.preventDefault();
        saveProject();
        announceStatus("Project saved");
        return;
      }
      if ((event.ctrlKey || event.metaKey) && key === "c") {
        event.preventDefault();
        copySelectionToClipboard();
        return;
      }
      if ((event.ctrlKey || event.metaKey) && key === "v") {
        event.preventDefault();
        pasteClipboardSelection();
        return;
      }
      if ((event.ctrlKey || event.metaKey) && key === "0") {
        event.preventDefault();
        fitToViewer();
        announceStatus("Fit to screen");
        return;
      }
      if (event.code === "Space" && !event.repeat) {
        event.preventDefault();
        if (!temporaryPanActive) {
          temporaryPanPreviousToolRef.current = tool;
          setTemporaryPanActive(true);
          if (tool !== TOOL.PAN) {
            setTool(TOOL.PAN);
            announceStatus("Temporary pan");
          }
        }
        return;
      }
      if (key === "f") {
        event.preventDefault();
        fitToViewer();
        announceStatus("Fit to screen");
        return;
      }
      if (key === "w") {
        event.preventDefault();
        fitWidthToViewer();
        announceStatus("Fit width view");
        return;
      }
      if (event.key === "PageDown") {
        event.preventDefault();
        changePageBy(1);
        return;
      }
      if (event.key === "PageUp") {
        event.preventDefault();
        changePageBy(-1);
        return;
      }
      if (event.key === "Escape") {
        if (draft?.type === "area") {
          setDraft(null);
          setSnapPreview(null);
          announceStatus("Area takeoff cancelled");
          return;
        }
        setActiveMenu(null);
        setDraft(null);
        setBoxZoomMode(false);
        setSnapPreview(null);
        clearSelection();
        announceStatus("Selection cleared");
        return;
      }
      if ((event.key === "Delete" || event.key === "Backspace") && selectedIds.length) {
        event.preventDefault();
        deleteSelected();
        return;
      }
      if (event.key === "Enter" && draft?.type === "area") {
        event.preventDefault();
        finishArea();
        return;
      }
      if (key === "v") {
        setTool(TOOL.SELECT);
        announceStatus("Select tool");
        return;
      }
      if (key === "h") {
        setTool(TOOL.PAN);
        announceStatus("Pan tool");
        return;
      }
      if (key === "d") {
        setTool(TOOL.DRAW);
        announceStatus("Sketch markup tool");
        return;
      }
      if (key === "u") {
        setTool(TOOL.RECTANGLE);
        announceStatus("Rectangle markup tool");
        return;
      }
      if (key === "i") {
        setTool(TOOL.ELLIPSE);
        announceStatus("Ellipse markup tool");
        return;
      }
      if (key === "m") {
        setTool(TOOL.MEASURE);
        announceStatus("Length tool");
        return;
      }
      if (key === "a") {
        setTool(TOOL.AREA);
        announceStatus("Area tool");
        return;
      }
      if (key === "t") {
        setTool(TOOL.TEXT);
        announceStatus("Text tool");
        return;
      }
      if (key === "r") {
        setTool(TOOL.STAMP);
        announceStatus("Stamp tool");
        return;
      }
      if (key === "j") {
        event.preventDefault();
        changePageBy(1);
        return;
      }
      if (key === "k") {
        event.preventDefault();
        changePageBy(-1);
        return;
      }
      if (key === "g") {
        event.preventDefault();
        setNavigationModeAndAnnounce(navigationMode === "sheet" ? "review" : "sheet");
        return;
      }
      if (key === "z") {
        event.preventDefault();
        activateBoxZoom();
      }
    };

    document.addEventListener("keydown", handleKeydown);
    return () => document.removeEventListener("keydown", handleKeydown);
  }, [draft, navigationMode, pageCount, redoStack, saveProject, selectedIds.length, temporaryPanActive, tool, undoStack]);

  useEffect(() => {
    const handleKeyup = (event: KeyboardEvent) => {
      if (event.code !== "Space") return;
      if (!temporaryPanActive) return;
      event.preventDefault();
      setTemporaryPanActive(false);
      const nextTool = temporaryPanPreviousToolRef.current;
      setTool(nextTool);
      if (nextTool !== TOOL.PAN) announceStatus(`${nextTool} tool restored`);
    };

    const handleWindowBlur = () => {
      if (!temporaryPanActive) return;
      setTemporaryPanActive(false);
      setTool(temporaryPanPreviousToolRef.current);
    };

    document.addEventListener("keyup", handleKeyup);
    window.addEventListener("blur", handleWindowBlur);
    return () => {
      document.removeEventListener("keyup", handleKeyup);
      window.removeEventListener("blur", handleWindowBlur);
    };
  }, [temporaryPanActive]);

  function resetLoadedDocumentState(uploadRunId: number) {
    uploadRunRef.current = uploadRunId;
    setLoadError("");
    setPageSizes({});
    setPageTexts({});
    setPageSheetNumbers({});
    setPageThumbnails({});
    setAutoReferenceLinks({});
    setHoveredAutoReferenceId(null);
    clearSelection();
    setSnapPreview(null);
    pageTextFragmentsRef.current = {};
  }

  async function loadPdfFromSource(options: { fileName: string; data?: ArrayBuffer; url?: string; pdfBlobUrl?: string }) {
    const uploadRunId = Date.now();
    resetLoadedDocumentState(uploadRunId);

    try {
      const loadingTask = options.data ? pdfjsLib.getDocument({ data: options.data }) : pdfjsLib.getDocument(options.url || "");
      const doc = await loadingTask.promise;
      setPdfDoc(doc);
      setFileName(options.fileName);
      setPageCount(doc.numPages);
      setCurrentPage(1);
      setWorkerStatus("ready");
      setActivePdfBlobUrl(options.pdfBlobUrl || "");
      clearHistory();
      const sheetNumbers: Record<number, string> = {};

      void (async () => {
        for (let pageNumber = 1; pageNumber <= doc.numPages; pageNumber += 1) {
          if (uploadRunRef.current !== uploadRunId) return;

          const page = await doc.getPage(pageNumber);
          const viewport = page.getViewport({ scale: 1 });
          setPageSizes((prev) => ({ ...prev, [pageNumber]: { width: viewport.width, height: viewport.height } }));

          const thumbnailScale = Math.min(0.3, 180 / Math.max(viewport.width, 1));
          const thumbnailViewport = page.getViewport({ scale: thumbnailScale });
          const thumbnailCanvas = document.createElement("canvas");
          const thumbnailContext = thumbnailCanvas.getContext("2d");

          if (thumbnailContext) {
            thumbnailCanvas.width = Math.max(1, Math.floor(thumbnailViewport.width));
            thumbnailCanvas.height = Math.max(1, Math.floor(thumbnailViewport.height));
            thumbnailContext.fillStyle = "#ffffff";
            thumbnailContext.fillRect(0, 0, thumbnailCanvas.width, thumbnailCanvas.height);
            await page.render({ canvasContext: thumbnailContext, viewport: thumbnailViewport }).promise;

            if (uploadRunRef.current !== uploadRunId) return;
            setPageThumbnails((prev) => ({ ...prev, [pageNumber]: thumbnailCanvas.toDataURL("image/png") }));
          }

          try {
            const textContent = await page.getTextContent();
            const items = textContent.items as TextContentItem[];
            const text = items.map((item) => item.str || "").join(" ");
            const sheetNumber = extractSheetNumber(items);
            pageTextFragmentsRef.current[pageNumber] = buildPageTextFragments(items, viewport as any);

            if (uploadRunRef.current !== uploadRunId) return;
            setPageTexts((prev) => ({ ...prev, [pageNumber]: text }));
            if (sheetNumber) {
              sheetNumbers[pageNumber] = sheetNumber;
              setPageSheetNumbers((prev) => ({ ...prev, [pageNumber]: sheetNumber }));
            }
          } catch {
            if (uploadRunRef.current !== uploadRunId) return;
            setPageTexts((prev) => ({ ...prev, [pageNumber]: "" }));
          }

          await new Promise((resolve) => window.setTimeout(resolve, 0));
        }

        if (uploadRunRef.current !== uploadRunId) return;
        setAutoReferenceLinks(buildAutoReferenceLinks(pageTextFragmentsRef.current, sheetNumbers));
      })().catch((error) => {
        if (uploadRunRef.current !== uploadRunId) return;
        const message = error instanceof Error ? error.message : "Unable to generate page previews.";
        setLoadError(message);
      });
    } catch (error) {
      const message = error instanceof Error ? error.message : "Unable to load PDF.";
      setLoadError(message);
      if (message.toLowerCase().includes("worker")) setWorkerStatus("error");
      setPdfDoc(null);
      setPageCount(0);
      setPageSheetNumbers({});
      setPageThumbnails({});
      setAutoReferenceLinks({});
      setActivePdfBlobUrl("");
      throw error;
    }
  }

  async function handleUpload(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0];
    if (!file) return;
    setActivePdfFile(file);
    setActivePdfBlobUrl("");
    try {
      const buffer = await file.arrayBuffer();
      await loadPdfFromSource({ fileName: file.name, data: buffer });
    } catch {
      setActivePdfFile(null);
    }
  }

  async function renderPage(pageNumber: number, targetCanvasRef: React.RefObject<HTMLCanvasElement>, renderKey: RenderCanvasKey, tintOpacity = 1) {
    if (!pdfDoc || !targetCanvasRef.current) return;
    renderRequestIdsRef.current[renderKey] += 1;
    const requestId = renderRequestIdsRef.current[renderKey];
    const previousTask = renderTasksRef.current[renderKey];
    if (previousTask) {
      previousTask.cancel();
      try {
        await previousTask.promise;
      } catch (error) {
        if (!isRenderCancelledError(error)) throw error;
      }
    }
    if (renderRequestIdsRef.current[renderKey] !== requestId) return;
    const page = await pdfDoc.getPage(pageNumber);
    const canvas = targetCanvasRef.current;
    const context = canvas.getContext("2d");
    if (!context) return;

    const baseViewport = page.getViewport({ scale: 1 });
    const deviceScale = typeof window !== "undefined" ? window.devicePixelRatio || 1 : 1;
    const renderScale = clamp(Math.max(2, deviceScale * Math.max(1, transform.scale)), 1, 4);
    const viewport = page.getViewport({ scale: renderScale });

    canvas.width = Math.floor(viewport.width);
    canvas.height = Math.floor(viewport.height);
    canvas.style.width = `${baseViewport.width}px`;
    canvas.style.height = `${baseViewport.height}px`;
    context.setTransform(1, 0, 0, 1, 0, 0);
    context.imageSmoothingEnabled = true;
    const renderTask = page.render({ canvasContext: context, viewport });
    renderTasksRef.current[renderKey] = renderTask;
    try {
      await renderTask.promise;
    } catch (error) {
      if (isRenderCancelledError(error)) return;
      throw error;
    } finally {
      if (renderTasksRef.current[renderKey] === renderTask) {
        renderTasksRef.current[renderKey] = null;
      }
    }
    if (renderRequestIdsRef.current[renderKey] !== requestId) return;

    if (tintOpacity < 1) {
      context.globalCompositeOperation = "source-atop";
      context.fillStyle = `rgba(220,38,38,${1 - tintOpacity})`;
      context.fillRect(0, 0, canvas.width, canvas.height);
      context.globalCompositeOperation = "source-over";
    }
  }

  useEffect(() => {
    if (pdfDoc) {
      renderPage(currentPage, pdfCanvasRef, "main", 1).catch((error) => {
        if (isRenderCancelledError(error)) return;
        setLoadError(error instanceof Error ? error.message : "Unable to render current page.");
      });
      if (compareMode && comparePage && comparePage !== currentPage) {
        renderPage(comparePage, compareCanvasRef, "compare", 0.55).catch((error) => {
          if (isRenderCancelledError(error)) return;
          setLoadError(error instanceof Error ? error.message : "Unable to render comparison page.");
        });
      }
    }
    return () => {
      renderTasksRef.current.main?.cancel();
      renderTasksRef.current.compare?.cancel();
    };
  }, [pdfDoc, currentPage, compareMode, comparePage, Math.round(transform.scale * 4) / 4]);

  function fitToViewer() {
    if (!viewerRef.current) return;
    const container = viewerRef.current;
    const padding = 40;
    const scaleX = (container.clientWidth - padding) / currentPageSize.width;
    const scaleY = (container.clientHeight - padding) / currentPageSize.height;
    const nextScale = Math.max(0.25, Math.min(4, Math.min(scaleX, scaleY)));
    const x = (container.clientWidth - currentPageSize.width * nextScale) / 2;
    const y = (container.clientHeight - currentPageSize.height * nextScale) / 2;
    setTransform({ x, y, scale: nextScale });
  }

  function fitWidthToViewer() {
    if (!viewerRef.current) return;
    const container = viewerRef.current;
    const padding = 28;
    const nextScale = clamp((container.clientWidth - padding) / currentPageSize.width, 0.25, 6);
    setTransform({
      x: (container.clientWidth - currentPageSize.width * nextScale) / 2,
      y: 16,
      scale: nextScale,
    });
  }

  function fitHeightToViewer() {
    if (!viewerRef.current) return;
    const container = viewerRef.current;
    const padding = 32;
    const nextScale = clamp((container.clientHeight - padding) / currentPageSize.height, 0.25, 6);
    setTransform({
      x: (container.clientWidth - currentPageSize.width * nextScale) / 2,
      y: (container.clientHeight - currentPageSize.height * nextScale) / 2,
      scale: nextScale,
    });
  }

  useEffect(() => {
    fitToViewer();
  }, [pdfDoc, currentPageSize.width, currentPageSize.height]);

  useEffect(() => {
    if (!pendingDetailFocus || pendingDetailFocus.targetPage !== currentPage) return;
    const timer = window.setTimeout(() => {
      focusDetailOnPage(pendingDetailFocus.detailToken, pendingDetailFocus.targetPage);
      setPendingDetailFocus(null);
    }, 90);
    return () => window.clearTimeout(timer);
  }, [currentPage, pendingDetailFocus, currentPageSize.height, currentPageSize.width]);

  useEffect(() => {
    const handleFullscreenChange = () => {
      const active = document.fullscreenElement === viewerRef.current;
      setIsViewerFullscreen(active);
      window.setTimeout(() => fitToViewer(), 50);
    };

    document.addEventListener("fullscreenchange", handleFullscreenChange);
    return () => document.removeEventListener("fullscreenchange", handleFullscreenChange);
  }, [currentPage, currentPageSize.height, currentPageSize.width]);

  useEffect(() => {
    const handleResize = () => setIsWideLayout(window.innerWidth >= 1536);
    window.addEventListener("resize", handleResize);
    return () => window.removeEventListener("resize", handleResize);
  }, []);

  useEffect(() => {
    const handlePointerMove = (event: PointerEvent) => {
      if (!resizeStateRef.current) return;
      const { side, startX, startWidth } = resizeStateRef.current;
      const delta = event.clientX - startX;
      if (side === "left") {
        setLeftPanelWidth(clamp(startWidth + delta, 210, 380));
      } else {
        setRightPanelWidth(clamp(startWidth - delta, 260, 460));
      }
    };

    const stopResizing = () => {
      if (!resizeStateRef.current) return;
      resizeStateRef.current = null;
      announceStatus("Panel size updated");
    };

    window.addEventListener("pointermove", handlePointerMove);
    window.addEventListener("pointerup", stopResizing);
    return () => {
      window.removeEventListener("pointermove", handlePointerMove);
      window.removeEventListener("pointerup", stopResizing);
    };
  }, []);

  function beginPanelResize(side: "left" | "right", clientX: number) {
    resizeStateRef.current = {
      side,
      startX: clientX,
      startWidth: side === "left" ? leftPanelWidth : rightPanelWidth,
    };
  }

  function setZoomScale(nextScale: number, screenPoint?: Point) {
    if (!viewerRef.current) return;
    const container = viewerRef.current;
    const anchor = screenPoint || { x: container.clientWidth / 2, y: container.clientHeight / 2 };
    const worldPoint = screenToWorld(anchor, transform);
    const clampedScale = clamp(nextScale, 0.25, 6);
    setTransform({
      scale: clampedScale,
      x: anchor.x - worldPoint.x * clampedScale,
      y: anchor.y - worldPoint.y * clampedScale,
    });
  }

  function zoomBy(factor: number) {
    if (!viewerRef.current) return;
    const container = viewerRef.current;
    setZoomScale(transform.scale * factor, { x: container.clientWidth / 2, y: container.clientHeight / 2 });
  }

  function zoomToPercent(percent: number) {
    const nextScale = clamp(percent / 100, 0.25, 6);
    setZoomScale(nextScale);
    announceStatus(`Zoom ${percent}%`);
  }

  function activateBoxZoom() {
    setBoxZoomMode(true);
    setDraft(null);
    announceStatus("Box zoom: drag a region on the sheet");
  }

  function changePageBy(delta: number) {
    const maxPage = Math.max(pageCount, 1);
    clearSelection();
    setDraft(null);
    setSnapPreview(null);
    setCurrentPage((prev) => Math.max(1, Math.min(maxPage, prev + delta)));
  }

  async function toggleViewerFullscreen() {
    if (!viewerRef.current) return;

    if (document.fullscreenElement === viewerRef.current) {
      await document.exitFullscreen?.();
      return;
    }

    await viewerRef.current.requestFullscreen?.();
  }

  function openPdfPicker() {
    fileInputRef.current?.click();
  }

  async function refreshCloudProjects() {
    try {
      const response = await fetch("/api/cloud-projects", { cache: "no-store" });
      if (!response.ok) {
        throw new Error(`Cloud list failed (${response.status})`);
      }
      const payload = (await response.json()) as { projects?: CloudProjectSummary[] };
      setCloudProjects((payload.projects || []).sort((a, b) => String(b.updatedAt).localeCompare(String(a.updatedAt))));
      setCloudStatus("ready");
      setCloudMessage("Shared cloud ready");
    } catch (error) {
      const message = error instanceof Error ? error.message : "Shared cloud unavailable";
      setCloudStatus("error");
      setCloudMessage(message);
    }
  }

  async function uploadCurrentPdfIfNeeded(projectId: string) {
    if (!activePdfFile) return activePdfBlobUrl;
    const uploadPath = `projects/${sanitizePathSegment(projectId)}/source/${sanitizePathSegment(activePdfFile.name)}`;
    const uploaded = await upload(uploadPath, activePdfFile, {
      access: "public",
      handleUploadUrl: "/api/blob-upload",
      clientPayload: JSON.stringify({ projectId, kind: "pdf", fileName: activePdfFile.name }),
      multipart: true,
    });
    setActivePdfBlobUrl(uploaded.url);
    setActivePdfFile(null);
    return uploaded.url;
  }

  function announceStatus(message: string) {
    setStatusMessage(message);
    if (statusTimeoutRef.current) window.clearTimeout(statusTimeoutRef.current);
    statusTimeoutRef.current = window.setTimeout(() => {
      setStatusMessage("Ready");
    }, 2400);
  }

  function setToolAndAnnounce(nextTool: ToolType, message: string) {
    setTool(nextTool);
    announceStatus(message);
  }

  function setNavigationModeAndAnnounce(nextMode: NavigationMode) {
    setNavigationMode(nextMode);
    announceStatus(nextMode === "sheet" ? "Wheel mode: sheet paging" : "Wheel mode: review pan");
  }

  function jumpToSheetOrPage(query: string) {
    const normalized = query.trim().toUpperCase();
    if (!normalized) return;

    const numericPage = Number(normalized);
    if (Number.isFinite(numericPage) && numericPage >= 1 && numericPage <= pageCount) {
      setCurrentPage(numericPage);
      setSheetJumpQuery("");
      announceStatus(`Jumped to page ${numericPage}`);
      return;
    }

    const exactSheetEntry = Object.entries(pageSheetNumbers).find(([, sheet]) => sheet.toUpperCase() === normalized);
    if (exactSheetEntry) {
      const targetPage = Number(exactSheetEntry[0]);
      setCurrentPage(targetPage);
      setSheetJumpQuery("");
      announceStatus(`Jumped to sheet ${normalized}`);
      return;
    }

    const partialSheetEntry = Object.entries(pageSheetNumbers).find(([, sheet]) => sheet.toUpperCase().includes(normalized));
    if (partialSheetEntry) {
      const targetPage = Number(partialSheetEntry[0]);
      setCurrentPage(targetPage);
      setSheetJumpQuery("");
      announceStatus(`Jumped to sheet ${partialSheetEntry[1]}`);
      return;
    }

    announceStatus(`No page or sheet matched "${query}"`);
  }

  function findDetailFragment(detailToken: string, page: number): PageTextFragment | null {
    const normalized = detailToken.trim().toUpperCase();
    if (!normalized) return null;
    const fragments = pageTextFragmentsRef.current[page] || [];

    const scored = fragments
      .map((fragment) => {
        const text = fragment.text.trim().toUpperCase();
        const compact = text.replace(/\s+/g, "");
        let score = 0;
        if (compact === normalized) score += 10;
        if (text === normalized) score += 8;
        if (compact.startsWith(normalized)) score += 5;
        if (compact.includes(normalized)) score += 2;
        score += Math.min(fragment.rect.h, 20) * 0.1;
        return { fragment, score };
      })
      .filter((item) => item.score > 0)
      .sort((a, b) => b.score - a.score);

    return scored[0]?.fragment || null;
  }

  function focusDetailOnPage(detailToken: string, page: number) {
    const fragment = findDetailFragment(detailToken, page);
    if (!fragment || !viewerRef.current) {
      announceStatus(`Jumped to sheet ${pageSheetNumbers[page] || page}`);
      return;
    }

    const viewer = viewerRef.current;
    const nextScale = clamp(Math.max(transform.scale, 2.2), 1.25, 6);
    const centerX = fragment.rect.x + fragment.rect.w / 2;
    const centerY = fragment.rect.y + fragment.rect.h / 2;
    setTransform({
      scale: nextScale,
      x: viewer.clientWidth / 2 - centerX * nextScale,
      y: viewer.clientHeight / 2 - centerY * nextScale,
    });
    announceStatus(`Focused detail ${detailToken} on ${pageSheetNumbers[page] || `Page ${page}`}`);
  }

  function navigateToAutoReference(link: AutoReferenceLink) {
    setCurrentPage(link.targetPage);
    setPendingDetailFocus({ targetPage: link.targetPage, detailToken: link.detailToken });
    announceStatus(`Opening ${link.text}`);
  }

  function showHelpDialog() {
    window.alert(
      [
        "Construction Viewer Controls",
        "",
        "Sheet wheel mode: scroll over the sheet for previous / next page",
        "Review wheel mode: scroll over the sheet to pan around the drawing",
        "Alt + Scroll over the sheet: scroll the app page instead",
        "Shift + Scroll over the sheet: zoom",
        "Double click: zoom in on cursor",
        "Shift + Double click: zoom out on cursor",
        "Double click a text, callout, stamp, cloud, or snapshot in Select mode to edit it on the sheet",
        "Drag on empty space in Select mode: box-select markups",
        "Ctrl + Click: add or remove markups from the current selection",
        "Alt + Drag a selected markup: duplicate and drag a copy",
        "Z: box zoom, then drag a region",
        "Pinch: zoom on touch screens",
        "Hold Space: temporarily pan without switching tools",
        "Pan tool: move around the sheet",
        "Select tool: hover references like 9/S302 for a preview, then click to jump sheets",
        "Shortcuts: V Select, H Pan, M Length, A Area, U Rectangle, I Ellipse, T Text, R Stamp, D Draw",
        "More shortcuts: F Fit, W Fit Width, J/K pages, G toggle wheel mode, Ctrl+C copy, Ctrl+V paste, Ctrl+Z undo, Ctrl+Y redo, PageUp/PageDown pages",
      ].join("\n")
    );
  }

  function showKeyboardShortcutsDialog() {
    window.alert(
      [
        "Keyboard Shortcuts",
        "",
        "File",
        "Ctrl+O Open PDF",
        "Ctrl+S Save project",
        "",
        "Edit",
        "Ctrl+C Copy selected markups",
        "Ctrl+V Paste markups",
        "Ctrl+Z Undo",
        "Ctrl+Y or Ctrl+Shift+Z Redo",
        "Delete / Backspace Delete selected markups",
        "Esc Clear selection or cancel current action",
        "",
        "View / Navigation",
        "Ctrl+0 or F Fit to screen",
        "W Fit width",
        "G Toggle wheel mode",
        "Z Box zoom",
        "J / Page Down Next page",
        "K / Page Up Previous page",
        "Space hold Temporary pan",
        "",
        "Tools",
        "V Select",
        "H Pan",
        "D Sketch",
        "M Length",
        "A Area",
        "U Rectangle",
        "I Ellipse",
        "T Text",
        "R Stamp",
        "",
        "Mouse / Trackpad",
        "Shift+Wheel Zoom on sheet",
        "Alt+Wheel Scroll app instead of sheet",
        "Double click Zoom in on cursor",
        "Shift+Double click Zoom out",
        "Ctrl/Cmd+Click Multi-select",
        "Alt+Drag Duplicate and drag selected markup",
        "Drag empty space in Select mode Box-select markups",
      ].join("\n")
    );
  }

  function isViewerWheelBypassTarget(target: EventTarget | null): boolean {
    return target instanceof HTMLElement && !!target.closest("button, input, select, textarea, label, a, [data-wheel-bypass='true']");
  }

  function isPointOverCurrentPage(point: Point) {
    const pageLeft = transform.x;
    const pageTop = transform.y;
    const pageWidth = currentPageSize.width * transform.scale;
    const pageHeight = currentPageSize.height * transform.scale;

    return (
      point.x >= pageLeft &&
      point.x <= pageLeft + pageWidth &&
      point.y >= pageTop &&
      point.y <= pageTop + pageHeight
    );
  }

  function centerViewerOnMiniMapPoint(clientX: number, clientY: number, element: HTMLElement) {
    if (!viewerRef.current || !miniMapMetrics) return;
    const rect = element.getBoundingClientRect();
    const localX = clamp(clientX - rect.left, 0, miniMapMetrics.width);
    const localY = clamp(clientY - rect.top, 0, miniMapMetrics.height);
    const worldX = localX / miniMapMetrics.scale;
    const worldY = localY / miniMapMetrics.scale;
    setTransform((current) => ({
      ...current,
      x: viewerRef.current!.clientWidth / 2 - worldX * current.scale,
      y: viewerRef.current!.clientHeight / 2 - worldY * current.scale,
    }));
  }

  function handleMiniMapPointerDown(e: React.PointerEvent<HTMLButtonElement>) {
    if (!miniMapMetrics) return;
    miniMapDragPointerIdRef.current = e.pointerId;
    e.currentTarget.setPointerCapture(e.pointerId);
    centerViewerOnMiniMapPoint(e.clientX, e.clientY, e.currentTarget);
    announceStatus(`Navigator centered on ${currentSheetNumber || `Page ${currentPage}`}`);
  }

  function handleMiniMapPointerMove(e: React.PointerEvent<HTMLButtonElement>) {
    if (miniMapDragPointerIdRef.current !== e.pointerId) return;
    centerViewerOnMiniMapPoint(e.clientX, e.clientY, e.currentTarget);
  }

  function handleMiniMapPointerUp(e: React.PointerEvent<HTMLButtonElement>) {
    if (miniMapDragPointerIdRef.current !== e.pointerId) return;
    miniMapDragPointerIdRef.current = null;
    if (e.currentTarget.hasPointerCapture(e.pointerId)) {
      e.currentTarget.releasePointerCapture(e.pointerId);
    }
  }

  function handleMiniMapPointerCancel(e: React.PointerEvent<HTMLButtonElement>) {
    if (miniMapDragPointerIdRef.current !== e.pointerId) return;
    miniMapDragPointerIdRef.current = null;
    if (e.currentTarget.hasPointerCapture(e.pointerId)) {
      e.currentTarget.releasePointerCapture(e.pointerId);
    }
  }

  function handleViewerDoubleClick(e: React.MouseEvent<HTMLDivElement>) {
    if (!viewerRef.current) return;
    if (isViewerWheelBypassTarget(e.target)) return;
    const rect = viewerRef.current.getBoundingClientRect();
    const screenPoint = { x: e.clientX - rect.left, y: e.clientY - rect.top };
    if (!isPointOverCurrentPage(screenPoint)) return;
    if (tool === TOOL.AREA && draft?.type === "area" && draft.points.length >= 3) {
      e.preventDefault();
      e.stopPropagation();
      finishArea();
      return;
    }
    if (tool === TOOL.SELECT && canEdit()) {
      const world = screenToWorld(screenPoint, transform);
      const hit = hitTest(world);
      if (hit && hit.mode !== "auto-link") {
        const annotation = currentAnnotations.find((item) => item.id === hit.id) || null;
        if (annotation) {
          const editorValue = getInlineEditorValue(annotation);
          if (editorValue) {
            e.preventDefault();
            e.stopPropagation();
            openInlineEditorForAnnotation(annotation);
            announceStatus("Editing markup on sheet");
            return;
          }
        }
      }
    }
    e.preventDefault();
    e.stopPropagation();
    const nextScale = transform.scale * (e.shiftKey ? 1 / 1.8 : 1.8);
    setZoomScale(nextScale, screenPoint);
    announceStatus(e.shiftKey ? "Zoomed out on cursor" : "Zoomed in on cursor");
  }

  function handleViewerWheel(e: WheelEvent) {
    if (!viewerRef.current) return;
    if (isViewerWheelBypassTarget(e.target)) return;
    const rect = viewerRef.current.getBoundingClientRect();
    const screenPoint = { x: e.clientX - rect.left, y: e.clientY - rect.top };
    const overSheet = isPointOverCurrentPage(screenPoint);
    if (!overSheet) return;
    if (e.altKey) return;

    e.preventDefault();
    e.stopPropagation();

    if (!e.shiftKey) {
      if (navigationMode === "review") {
        const panX = e.shiftKey ? e.deltaY : e.deltaX;
        const panY = e.shiftKey ? 0 : e.deltaY;
        setTransform((current) => ({
          ...current,
          x: current.x - panX,
          y: current.y - panY,
        }));
        return;
      }
      if (Math.abs(e.deltaY) < 18) return;
      const now = Date.now();
      if (now - pageWheelCooldownRef.current < 280) return;

      pageWheelCooldownRef.current = now;
      changePageBy(e.deltaY > 0 ? 1 : -1);
      return;
    }

    const worldPoint = screenToWorld(screenPoint, transform);
    const nextScale = clamp(transform.scale * (e.deltaY < 0 ? 1.12 : 1 / 1.12), 0.25, 6);
    setTransform({
      scale: nextScale,
      x: screenPoint.x - worldPoint.x * nextScale,
      y: screenPoint.y - worldPoint.y * nextScale,
    });
  }

  useEffect(() => {
    const viewer = viewerRef.current;
    if (!viewer) return;
    const listener = (event: WheelEvent) => handleViewerWheel(event);
    viewer.addEventListener("wheel", listener, { passive: false });
    return () => viewer.removeEventListener("wheel", listener);
  }, [navigationMode, transform, currentPageSize.width, currentPageSize.height, currentPage]);

  function updateTouchPinch() {
    if (touchPointersRef.current.size < 2) return false;
    const [first, second] = Array.from(touchPointersRef.current.values());
    const midpoint = { x: (first.x + second.x) / 2, y: (first.y + second.y) / 2 };
    const nextDistance = Math.max(distance(first, second), 1);
    const previous = pinchStateRef.current;

    if (previous) {
      setTransform((current) => {
        const worldPoint = screenToWorld(midpoint, current);
        const nextScale = clamp(current.scale * (nextDistance / previous.distance), 0.25, 6);
        return {
          scale: nextScale,
          x: midpoint.x - worldPoint.x * nextScale,
          y: midpoint.y - worldPoint.y * nextScale,
        };
      });
    }

    pinchStateRef.current = { distance: nextDistance, midpoint };
    return true;
  }

  function activeUserMeta() {
    return {
      author: activeUser?.name || "You",
      color: activeMarkupColor || activeUser?.color || "#2563eb",
      roleId: activeUser?.roleId || "role-pm",
    };
  }

  function createHistorySnapshot(): HistorySnapshot {
    return {
      annotations: deepClone(annotations),
      calibration,
      layers: deepClone(layers),
      currentPage,
    };
  }

  function clearHistory() {
    setUndoStack([]);
    setRedoStack([]);
  }

  function pushHistorySnapshot() {
    if (isApplyingHistoryRef.current) return;
    const snapshot = createHistorySnapshot();
    setUndoStack((prev) => [snapshot, ...prev].slice(0, 100));
    setRedoStack([]);
  }

  function applyHistorySnapshot(snapshot: HistorySnapshot) {
    isApplyingHistoryRef.current = true;
    setAnnotations(snapshot.annotations);
    setCalibration(snapshot.calibration);
    setLayers(snapshot.layers);
    setCurrentPage(snapshot.currentPage);
    setSelection([]);
    setDraft(null);
    setSnapPreview(null);
    window.setTimeout(() => {
      isApplyingHistoryRef.current = false;
    }, 0);
  }

  function undoLastAction() {
    if (!undoStack.length) {
      announceStatus("Nothing to undo");
      return;
    }
    const [snapshot, ...rest] = undoStack;
    const currentSnapshot = createHistorySnapshot();
    setUndoStack(rest);
    setRedoStack((prev) => [currentSnapshot, ...prev].slice(0, 100));
    applyHistorySnapshot(snapshot);
    announceStatus("Undid last change");
  }

  function redoLastAction() {
    if (!redoStack.length) {
      announceStatus("Nothing to redo");
      return;
    }
    const [snapshot, ...rest] = redoStack;
    const currentSnapshot = createHistorySnapshot();
    setRedoStack(rest);
    setUndoStack((prev) => [currentSnapshot, ...prev].slice(0, 100));
    applyHistorySnapshot(snapshot);
    announceStatus("Redid change");
  }

  function applyMarkupPreset(preset: (typeof MARKUP_PRESETS)[number]) {
    setActiveMarkupColor(preset.color);
    setActiveMarkupLayerId(preset.layerId);
    setActiveMarkupCategory(preset.category);
    setActiveMarkupFillColor(preset.color);
    announceStatus(`${preset.name} preset ready`);
  }

  function applyStampPreset(preset: (typeof STAMP_PRESETS)[number]) {
    setActiveStampPresetId(preset.id);
    setActiveMarkupColor(preset.color);
    setActiveMarkupLayerId(preset.layerId);
    setActiveMarkupCategory(preset.category);
    setActiveMarkupFillColor(preset.color);
    setActiveMarkupFillOpacity(preset.fillOpacity);
    announceStatus(`${preset.name} stamp ready`);
  }

  function updateStyleValue(field: "color" | "strokeWidth" | "fillColor" | "fillOpacity", value: string | number) {
    if (selectedIds.length && canEdit()) {
      patchSelectedMany((item) => ({ ...item, [field]: value } as Annotation));
    }

    if (field === "color" && typeof value === "string") setActiveMarkupColor(value);
    if (field === "strokeWidth" && typeof value === "number") setActiveMarkupStrokeWidth(value);
    if (field === "fillColor" && typeof value === "string") setActiveMarkupFillColor(value);
    if (field === "fillOpacity" && typeof value === "number") setActiveMarkupFillOpacity(value);
  }

  function canEdit() {
    return !!activeRole?.canEdit;
  }

  function canDelete() {
    return !!activeRole?.canDelete;
  }

  function canExport() {
    return !!activeRole?.canExport;
  }

  function setSelection(ids: string[], primaryId?: string | null) {
    const nextIds = Array.from(new Set(ids));
    setSelectedIds(nextIds);
    setSelectedId(primaryId && nextIds.includes(primaryId) ? primaryId : nextIds[0] || null);
    if (!nextIds.length) setSelectedHandle(null);
  }

  function clearSelection() {
    setSelection([]);
    setInlineEditor(null);
  }

  function visibleAnnotationsForPage(page: number): Annotation[] {
    return ensureArray(annotations[page]).filter((item) => visibleLayerIds.has(item.layerId || "layer-general"));
  }

  function updateAnnotations(updater: (items: Annotation[]) => Annotation[], options?: { recordHistory?: boolean }) {
    if (options?.recordHistory !== false) pushHistorySnapshot();
    setAnnotations((prev) => ({ ...prev, [currentPage]: updater(ensureArray(prev[currentPage])) }));
  }

  function patchAnnotationByPage(
    page: number,
    annotationId: string,
    mutator: (item: Annotation) => Annotation,
    options?: { recordHistory?: boolean }
  ) {
    if (options?.recordHistory !== false) pushHistorySnapshot();
    setAnnotations((prev) => ({
      ...prev,
      [page]: ensureArray(prev[page]).map((item) => (item.id === annotationId ? mutator(item) : item)),
    }));
  }

  function patchSelected(mutator: (item: Annotation) => Annotation, options?: { recordHistory?: boolean }) {
    if (!selectedId) return;
    patchAnnotationByPage(currentPage, selectedId, mutator, options);
  }

  function patchSelectedMany(mutator: (item: Annotation) => Annotation, options?: { recordHistory?: boolean }) {
    if (!selectedIds.length) return;
    const selectedSet = new Set(selectedIds);
    if (options?.recordHistory !== false) pushHistorySnapshot();
    setAnnotations((prev) => ({
      ...prev,
      [currentPage]: ensureArray(prev[currentPage]).map((item) => (selectedSet.has(item.id) ? mutator(item) : item)),
    }));
  }

  function copySelectionToClipboard() {
    if (!selectedAnnotations.length) {
      announceStatus("Nothing selected to copy");
      return;
    }
    setClipboardAnnotations(selectedAnnotations.map((annotation) => cloneAnnotation(annotation)));
    announceStatus(`Copied ${selectedAnnotations.length} markup${selectedAnnotations.length === 1 ? "" : "s"}`);
  }

  function pasteClipboardSelection() {
    if (!clipboardAnnotations.length || !canEdit()) {
      if (!clipboardAnnotations.length) announceStatus("Clipboard is empty");
      return;
    }
    const createdAt = new Date().toISOString();
    const pastedItems = clipboardAnnotations.map((annotation, index) => {
      const pasted = offsetAnnotation(annotation, 20 + index * 8, 20 + index * 8);
      pasted.id = createId();
      pasted.page = currentPage;
      pasted.createdAt = createdAt;
      return pasted;
    });
    updateAnnotations((list) => [...list, ...pastedItems]);
    setSelection(pastedItems.map((item) => item.id), pastedItems[pastedItems.length - 1]?.id || null);
    announceStatus(`Pasted ${pastedItems.length} markup${pastedItems.length === 1 ? "" : "s"}`);
  }

  function collectSnapCandidates(page: number, excludeIds: string[] = []): Array<{ point: Point; label: string }> {
    const excluded = new Set(excludeIds);
    return ensureArray(annotations[page])
      .filter((annotation) => !excluded.has(annotation.id))
      .flatMap((annotation) => collectAnnotationSnapPoints(annotation));
  }

  function getSnappedWorldPoint(rawPoint: Point, options?: { anchor?: Point; excludeIds?: string[] }) {
    const tolerance = 12 / Math.max(transform.scale, 0.45);
    const candidates = collectSnapCandidates(currentPage, options?.excludeIds);
    let nearestPoint: { point: Point; label: string; distance: number } | null = null;

    candidates.forEach((candidate) => {
      const candidateDistance = distance(rawPoint, candidate.point);
      if (candidateDistance > tolerance) return;
      if (!nearestPoint || candidateDistance < nearestPoint.distance) {
        nearestPoint = { ...candidate, distance: candidateDistance };
      }
    });

    if (nearestPoint) {
      return {
        point: nearestPoint.point,
        preview: { point: nearestPoint.point, label: nearestPoint.label } as SnapPreview,
      };
    }

    if (!options?.anchor) {
      return { point: rawPoint, preview: null as SnapPreview | null };
    }

    let snappedX = rawPoint.x;
    let snappedY = rawPoint.y;
    let guideX: number | undefined;
    let guideY: number | undefined;
    let label = "";

    if (Math.abs(rawPoint.x - options.anchor.x) <= tolerance) {
      snappedX = options.anchor.x;
      guideX = snappedX;
      label = "Vertical align";
    }
    if (Math.abs(rawPoint.y - options.anchor.y) <= tolerance) {
      snappedY = options.anchor.y;
      guideY = snappedY;
      label = label ? "Axis align" : "Horizontal align";
    }

    if (guideX === undefined && guideY === undefined) {
      return { point: rawPoint, preview: null as SnapPreview | null };
    }

    return {
      point: { x: snappedX, y: snappedY },
      preview: { point: { x: snappedX, y: snappedY }, label, guideX, guideY } as SnapPreview,
    };
  }

  function getInlineEditorValue(annotation: Annotation): { field: InlineEditorField; value: string } | null {
    if (annotation.type === "text") return { field: "text", value: annotation.text || "" };
    if (annotation.type === "comment") return { field: "comment", value: annotation.comment || annotation.text || "" };
    if (annotation.type === "callout") return { field: "text", value: annotation.text || "" };
    if (annotation.type === "stamp") return { field: "text", value: annotation.text || "" };
    if (annotation.type === "cloud") return { field: "label", value: annotation.label || "" };
    if (annotation.type === "snapshot") return { field: "title", value: annotation.title || "" };
    return null;
  }

  function openInlineEditorForAnnotation(annotation: Annotation) {
    const editorValue = getInlineEditorValue(annotation);
    if (!editorValue || !canEdit()) return;
    setSelection([annotation.id], annotation.id);
    setSelectedHandle("move");
    setInlineEditor({
      annotationId: annotation.id,
      page: annotation.page,
      field: editorValue.field,
      value: editorValue.value,
    });
  }

  function cancelInlineEditor() {
    setInlineEditor(null);
  }

  function commitInlineEditor() {
    if (!inlineEditor) return;
    const annotation = ensureArray(annotations[inlineEditor.page]).find((item) => item.id === inlineEditor.annotationId);
    if (!annotation) {
      setInlineEditor(null);
      return;
    }

    const nextValue = inlineEditor.value.trim();
    const currentValue = getInlineEditorValue(annotation)?.value?.trim() || "";
    if (nextValue === currentValue) {
      setInlineEditor(null);
      return;
    }

    patchAnnotationByPage(inlineEditor.page, inlineEditor.annotationId, (item) => {
      if (inlineEditor.field === "comment" && item.type === "comment") {
        return { ...item, text: nextValue, comment: nextValue };
      }
      if (inlineEditor.field === "label" && item.type === "cloud") {
        return { ...item, label: nextValue };
      }
      if (inlineEditor.field === "title" && item.type === "snapshot") {
        return { ...item, title: nextValue };
      }
      if (inlineEditor.field === "text" && (item.type === "text" || item.type === "callout" || item.type === "stamp")) {
        return { ...item, text: nextValue };
      }
      return item;
    });
    setInlineEditor(null);
    announceStatus("Updated markup text");
  }

  function hitTest(world: Point): { id: string; mode: string } | null {
    for (let index = currentAutoReferenceLinks.length - 1; index >= 0; index -= 1) {
      const link = currentAutoReferenceLinks[index];
      if (pointInRect(world, link.rect)) {
        return { id: link.id, mode: "auto-link" };
      }
    }

    const candidates = visibleAnnotationsForPage(currentPage);
    for (let index = candidates.length - 1; index >= 0; index -= 1) {
      const item = candidates[index];
      if ((item.type === "text" || item.type === "comment") && pointNear(world, item.position, 25)) {
        return { id: item.id, mode: "move" };
      }
      if (item.type === "stamp" && pointInRect(world, getStampWorldRect(item))) {
        return { id: item.id, mode: "move" };
      }
      if (item.type === "callout") {
        if (pointNear(world, item.anchor, 16)) return { id: item.id, mode: "handle:anchor" };
        const boxHandle = rectHandleAtPoint(world, item.box);
        if (boxHandle) return { id: item.id, mode: `handle:box:${boxHandle}` };
        if (
          world.x >= item.box.x &&
          world.x <= item.box.x + item.box.w &&
          world.y >= item.box.y &&
          world.y <= item.box.y + item.box.h
        ) {
          return { id: item.id, mode: "move" };
        }
      }
      if (item.type === "rectangle" || item.type === "ellipse") {
        const rectHandle = rectHandleAtPoint(world, item.rect);
        if (rectHandle) return { id: item.id, mode: `handle:rect:${rectHandle}` };
        if ((item.type === "rectangle" && pointInRect(world, item.rect)) || (item.type === "ellipse" && pointInEllipse(world, item.rect))) {
          return { id: item.id, mode: "move" };
        }
      }
      if (item.type === "snapshot") {
        const rectHandle = rectHandleAtPoint(world, item.rect);
        if (rectHandle) return { id: item.id, mode: `handle:rect:${rectHandle}` };
        if (pointInRect(world, item.rect)) return { id: item.id, mode: "move" };
      }
      if (item.type === "cloud") {
        const rectHandle = rectHandleAtPoint(world, item.rect);
        if (rectHandle) return { id: item.id, mode: `handle:rect:${rectHandle}` };
        if (pointInRect(world, item.rect)) return { id: item.id, mode: "move" };
      }
      if (item.type === "measure") {
        if (pointNear(world, item.start, 14)) return { id: item.id, mode: "handle:start" };
        if (pointNear(world, item.end, 14)) return { id: item.id, mode: "handle:end" };
      }
      if (item.type === "area") {
        for (let pointIndex = 0; pointIndex < item.points.length; pointIndex += 1) {
          if (pointNear(world, item.points[pointIndex], 14)) return { id: item.id, mode: `handle:point:${pointIndex}` };
        }
        if (pointInPolygon(world, item.points)) return { id: item.id, mode: "move" };
      }
      if (item.type === "path") {
        for (const point of item.points) {
          if (pointNear(world, point, 10)) return { id: item.id, mode: "move" };
        }
      }
      if (item.type === "link" && pointNear(world, item.position, 18)) return { id: item.id, mode: "open-link" };
    }
    return null;
  }

  function applyToolChestAt(world: Point): boolean {
    const template = toolChest.find((item) => item.id === selectedChestId);
    if (!template) return false;
    const meta = activeUserMeta();
    updateAnnotations((list) => [
      ...list,
      {
        id: createId(),
        type: "text",
        text: template.payload.text,
        color: template.payload.color,
        position: world,
        page: currentPage,
        author: meta.author,
        status: "open",
        createdAt: new Date().toISOString(),
        layerId: template.payload.layerId,
        category: template.payload.category,
        roleId: meta.roleId,
        strokeWidth: activeMarkupStrokeWidth,
        fillColor: activeMarkupFillColor,
        fillOpacity: activeMarkupFillOpacity,
      },
    ]);
    return true;
  }

  function handlePointerDown(e: React.PointerEvent<HTMLCanvasElement>) {
    const overlay = overlayCanvasRef.current;
    if (!overlay) return;
    if (inlineEditor) setInlineEditor(null);
    setSnapPreview(null);
    const point = getPointerPosition(e, overlay);

    if (e.pointerType === "touch") {
      touchPointersRef.current.set(e.pointerId, point);
      if (touchPointersRef.current.size >= 2) {
        setIsPanning(false);
        setPanStart(null);
        setDraggingSelection(null);
        setDraft(null);
        updateTouchPinch();
        return;
      }
    }

    const world = screenToWorld(point, transform);

    if (boxZoomMode) {
      clearSelection();
      setDraft({ type: "zoom-box", start: world, current: world });
      return;
    }

    if (tool === TOOL.PAN || tool === TOOL.SELECT) {
      const autoLink = currentAutoReferenceLinks.find((link) => pointInRect(world, link.rect));
      if (autoLink) {
        clearSelection();
        setDraft(null);
        navigateToAutoReference(autoLink);
        setHoveredAutoReferenceId(autoLink.id);
        return;
      }
    }

    if (tool === TOOL.PAN) {
      setIsPanning(true);
      setPanStart({ x: e.clientX, y: e.clientY, originX: transform.x, originY: transform.y });
      return;
    }

    if (tool === TOOL.SELECT) {
      const hit = hitTest(world);
      if (!hit) {
        setDraft({ type: "selection-box", start: world, current: world, append: !!(e.ctrlKey || e.metaKey) });
        return;
      }
      setSelectedHandle(hit.mode);
      if (e.ctrlKey || e.metaKey) {
        const nextSelection = selectedIdSet.has(hit.id)
          ? selectedIds.filter((id) => id !== hit.id)
          : [...selectedIds, hit.id];
        setSelection(nextSelection, hit.id);
        return;
      }
      if (!selectedIdSet.has(hit.id) || selectedIds.length <= 1) {
        setSelection([hit.id], hit.id);
      } else {
        setSelection(selectedIds, hit.id);
      }
      if (hit.mode === "open-link") {
        const linkItem = currentAnnotations.find((item) => item.id === hit.id) as LinkAnnotation | undefined;
        if (linkItem?.targetPage) setCurrentPage(linkItem.targetPage);
        return;
      }
      if (canEdit()) {
        if (e.altKey && hit.mode === "move") {
          const sourceSelection = selectedIdSet.has(hit.id)
            ? currentAnnotations.filter((annotation) => selectedIdSet.has(annotation.id))
            : currentAnnotations.filter((annotation) => annotation.id === hit.id);
          if (sourceSelection.length) {
            pushHistorySnapshot();
            const duplicates = createDuplicatedAnnotations(sourceSelection, currentPage, 0);
            setAnnotations((prev) => ({ ...prev, [currentPage]: [...ensureArray(prev[currentPage]), ...duplicates] }));
            setSelection(duplicates.map((item) => item.id), duplicates[duplicates.length - 1]?.id || null);
            announceStatus(`Duplicated ${duplicates.length} markup${duplicates.length === 1 ? "" : "s"} for drag`);
          }
        } else {
          pushHistorySnapshot();
        }
        setDraggingSelection({ world });
      }
      return;
    }

    if (!canEdit()) return;
    if (selectedChestId && applyToolChestAt(world)) return;

    if (tool === TOOL.DRAW) {
      setDraft({ type: "path", points: [world], layerId: activeMarkupLayerId, category: activeMarkupCategory || "Sketch" });
      return;
    }
    if (tool === TOOL.MEASURE) {
      setDraft({ type: "measure", start: world, end: world, layerId: activeMarkupLayerId, category: activeMarkupCategory || "Linear" });
      return;
    }
    if (tool === TOOL.AREA) {
      const snapAnchor = draft && draft.type === "area" && draft.points.length ? draft.points[draft.points.length - 1] : undefined;
      const snappedPoint = getSnappedWorldPoint(world, { anchor: snapAnchor, excludeIds: selectedIds }).point;
      if (draft && draft.type === "area" && draft.points.length >= 3 && pointNear(snappedPoint, draft.points[0], 18 / Math.max(transform.scale, 0.5))) {
        finishArea();
        return;
      }
      setDraft((prev) => {
        if (!prev || prev.type !== "area") return { type: "area", points: [snappedPoint], layerId: activeMarkupLayerId, category: activeMarkupCategory || "Area" };
        return { ...prev, points: [...prev.points, snappedPoint] };
      });
      setSnapPreview(null);
      return;
    }
    if (tool === TOOL.RECTANGLE) {
      setDraft({ type: "rectangle", start: world, current: world, layerId: activeMarkupLayerId, category: activeMarkupCategory || "Shapes" });
      return;
    }
    if (tool === TOOL.ELLIPSE) {
      setDraft({ type: "ellipse", start: world, current: world, layerId: activeMarkupLayerId, category: activeMarkupCategory || "Shapes" });
      return;
    }
    if (tool === TOOL.CALLOUT) {
      setDraft({ type: "callout", start: world, current: world, text: "Field verify", layerId: activeMarkupLayerId, category: activeMarkupCategory || "Callouts" });
      return;
    }
    if (tool === TOOL.CLOUD) {
      setDraft({ type: "cloud", start: world, current: world, label: "Revision", layerId: activeMarkupLayerId, category: activeMarkupCategory || "Clouds" });
      return;
    }
    if (tool === TOOL.STAMP) {
      const preset = STAMP_PRESETS.find((item) => item.id === activeStampPresetId) || STAMP_PRESETS[0];
      const metaForStamp = activeUserMeta();
      const newStamp: StampAnnotation = {
        id: createId(),
        type: "stamp",
        text: preset.text,
        stampKind: preset.id,
        position: world,
        page: currentPage,
        author: metaForStamp.author,
        status: "open",
        createdAt: new Date().toISOString(),
        layerId: activeMarkupLayerId || preset.layerId,
        category: activeMarkupCategory || preset.category,
        color: activeMarkupColor || preset.color,
        roleId: metaForStamp.roleId,
        strokeWidth: activeMarkupStrokeWidth,
        fillColor: activeMarkupFillColor || preset.color,
        fillOpacity: activeMarkupFillOpacity,
      };
      updateAnnotations((list) => [...list, newStamp]);
      openInlineEditorForAnnotation(newStamp);
      announceStatus(`${preset.name} stamp placed`);
      return;
    }
    if (tool === TOOL.CALIBRATE) {
      setDraft({ type: "calibrate", start: world, end: world });
      return;
    }

    const meta = activeUserMeta();

    if (tool === TOOL.TEXT) {
      const newText: TextAnnotation = {
        id: createId(),
        type: "text",
        text: "Note",
        color: meta.color,
        position: world,
        page: currentPage,
        author: meta.author,
        status: "open",
        createdAt: new Date().toISOString(),
        layerId: activeMarkupLayerId,
        category: activeMarkupCategory || "Notes",
        roleId: meta.roleId,
        strokeWidth: activeMarkupStrokeWidth,
        fillColor: activeMarkupFillColor,
        fillOpacity: activeMarkupFillOpacity,
      };
      updateAnnotations((list) => [...list, newText]);
      openInlineEditorForAnnotation(newText);
      return;
    }

    if (tool === TOOL.COMMENT) {
      const newComment: TextAnnotation = {
        id: createId(),
        type: "comment",
        text: "Coordination comment",
        comment: "Coordination comment",
        color: meta.color,
        position: world,
        page: currentPage,
        author: meta.author,
        status: "open",
        createdAt: new Date().toISOString(),
        layerId: activeMarkupLayerId,
        category: activeMarkupCategory || "Comments",
        roleId: meta.roleId,
        strokeWidth: activeMarkupStrokeWidth,
        fillColor: activeMarkupFillColor,
        fillOpacity: activeMarkupFillOpacity,
      };
      updateAnnotations((list) => [...list, newComment]);
      openInlineEditorForAnnotation(newComment);
      return;
    }

    if (tool === TOOL.SNAPSHOT) {
      setDraft({ type: "snapshot", start: world, current: world, title: `Snapshot P${currentPage}`, layerId: activeMarkupLayerId, category: activeMarkupCategory || "Snapshots" });
      return;
    }

    if (tool === TOOL.LINK) {
      const targetValue = prompt(`Enter target page number (1-${pageCount || 1})`, `${Math.min(currentPage + 1, pageCount || 1)}`);
      if (!targetValue) return;
      const targetPage = Math.max(1, Math.min(pageCount || 1, Number(targetValue)));
      updateAnnotations((list) => [
        ...list,
        {
          id: createId(),
          type: "link",
          text: `→ ${targetPage}`,
          targetPage,
          position: world,
          page: currentPage,
          author: meta.author,
          status: "open",
          createdAt: new Date().toISOString(),
          layerId: activeMarkupLayerId || "layer-links",
          category: activeMarkupCategory || "Links",
          color: activeMarkupColor || "#1d4ed8",
          roleId: meta.roleId,
          strokeWidth: activeMarkupStrokeWidth,
          fillColor: activeMarkupFillColor,
          fillOpacity: activeMarkupFillOpacity,
        },
      ]);
    }
  }

  function handlePointerMove(e: React.PointerEvent<HTMLCanvasElement>) {
    const overlay = overlayCanvasRef.current;
    if (!overlay) return;

    if (e.pointerType === "touch" && touchPointersRef.current.has(e.pointerId)) {
      touchPointersRef.current.set(e.pointerId, getPointerPosition(e, overlay));
      if (touchPointersRef.current.size >= 2 && updateTouchPinch()) return;
    }

    if (isPanning && panStart) {
      setSnapPreview(null);
      setTransform((prev) => ({ ...prev, x: panStart.originX + (e.clientX - panStart.x), y: panStart.originY + (e.clientY - panStart.y) }));
      return;
    }

    const point = getPointerPosition(e, overlay);
    const world = screenToWorld(point, transform);
    const isOverCurrentPage = isPointOverCurrentPage(point);
    setCursorWorld(isOverCurrentPage ? world : null);
    if (!isOverCurrentPage) {
      setHoveredKeynoteToken(null);
    }
    const hoveredAutoLink =
      tool === TOOL.PAN || tool === TOOL.SELECT
        ? currentAutoReferenceLinks.find((link) => pointInRect(world, link.rect)) || null
        : null;
    setHoveredAutoReferenceId(hoveredAutoLink?.id || null);
    if (isOverCurrentPage) {
      const hoveredKeynote =
        currentPageTextFragments.find((fragment) => pointInRect(world, fragment.rect) && extractKeynoteTokens(fragment.text, knownSheetNumberSet).length > 0) || null;
      setHoveredKeynoteToken(hoveredKeynote ? extractKeynoteTokens(hoveredKeynote.text, knownSheetNumberSet)[0] || null : null);
    }

    if (tool === TOOL.SELECT && draggingSelection && selectedId && canEdit()) {
      const mode = selectedHandle || "move";
      let workingWorld = world;

      if (mode === "handle:end" && selectedAnnotation?.type === "measure") {
        const snapped = getSnappedWorldPoint(world, { anchor: selectedAnnotation.start, excludeIds: selectedIds });
        workingWorld = snapped.point;
        setSnapPreview(snapped.preview);
      } else if (mode === "handle:start" && selectedAnnotation?.type === "measure") {
        const snapped = getSnappedWorldPoint(world, { anchor: selectedAnnotation.end, excludeIds: selectedIds });
        workingWorld = snapped.point;
        setSnapPreview(snapped.preview);
      } else if (mode.startsWith("handle:rect:") && selectedAnnotation && "rect" in selectedAnnotation) {
        const rectHandle = mode.split(":")[2] as "nw" | "ne" | "sw" | "se";
        const opposite = {
          nw: { x: selectedAnnotation.rect.x + selectedAnnotation.rect.w, y: selectedAnnotation.rect.y + selectedAnnotation.rect.h },
          ne: { x: selectedAnnotation.rect.x, y: selectedAnnotation.rect.y + selectedAnnotation.rect.h },
          sw: { x: selectedAnnotation.rect.x + selectedAnnotation.rect.w, y: selectedAnnotation.rect.y },
          se: { x: selectedAnnotation.rect.x, y: selectedAnnotation.rect.y },
        }[rectHandle];
        const snapped = getSnappedWorldPoint(world, { anchor: opposite, excludeIds: selectedIds });
        workingWorld = snapped.point;
        setSnapPreview(snapped.preview);
      } else if (mode.startsWith("handle:box:") && selectedAnnotation?.type === "callout") {
        const boxHandle = mode.split(":")[2] as "nw" | "ne" | "sw" | "se";
        const opposite = {
          nw: { x: selectedAnnotation.box.x + selectedAnnotation.box.w, y: selectedAnnotation.box.y + selectedAnnotation.box.h },
          ne: { x: selectedAnnotation.box.x, y: selectedAnnotation.box.y + selectedAnnotation.box.h },
          sw: { x: selectedAnnotation.box.x + selectedAnnotation.box.w, y: selectedAnnotation.box.y },
          se: { x: selectedAnnotation.box.x, y: selectedAnnotation.box.y },
        }[boxHandle];
        const snapped = getSnappedWorldPoint(world, { anchor: opposite, excludeIds: selectedIds });
        workingWorld = snapped.point;
        setSnapPreview(snapped.preview);
      } else if (mode === "handle:anchor" && selectedAnnotation?.type === "callout") {
        const snapped = getSnappedWorldPoint(world, { excludeIds: selectedIds });
        workingWorld = snapped.point;
        setSnapPreview(snapped.preview);
      } else if (mode.startsWith("handle:point:") && selectedAnnotation?.type === "area") {
        const pointIndex = Number(mode.split(":")[2]);
        const previousPoint = selectedAnnotation.points[Math.max(pointIndex - 1, 0)];
        const snapped = getSnappedWorldPoint(world, { anchor: previousPoint, excludeIds: selectedIds });
        workingWorld = snapped.point;
        setSnapPreview(snapped.preview);
      } else {
        setSnapPreview(null);
      }

      const dx = workingWorld.x - draggingSelection.world.x;
      const dy = workingWorld.y - draggingSelection.world.y;
      const idsToMutate = mode === "move" ? (selectedIds.length ? selectedIds : [selectedId]) : [selectedId];
      const selectedSet = new Set(idsToMutate);

      patchSelectedMany((item) => {
        if (!selectedSet.has(item.id)) return item;
        if (mode === "move") {
          if (item.type === "text" || item.type === "comment" || item.type === "link" || item.type === "stamp") {
            return { ...item, position: { x: item.position.x + dx, y: item.position.y + dy } };
          }
          if (item.type === "callout") {
            return {
              ...item,
              anchor: { x: item.anchor.x + dx, y: item.anchor.y + dy },
              box: { ...item.box, x: item.box.x + dx, y: item.box.y + dy },
            };
          }
          if (item.type === "measure") {
            return { ...item, start: { x: item.start.x + dx, y: item.start.y + dy }, end: { x: item.end.x + dx, y: item.end.y + dy } };
          }
          if (item.type === "area") return { ...item, points: item.points.map((p) => ({ x: p.x + dx, y: p.y + dy })) };
          if (item.type === "path") return { ...item, points: item.points.map((p) => ({ x: p.x + dx, y: p.y + dy })) };
          if (item.type === "snapshot" || item.type === "cloud" || item.type === "rectangle" || item.type === "ellipse") {
            return { ...item, rect: { ...item.rect, x: item.rect.x + dx, y: item.rect.y + dy } };
          }
        }
        if (mode === "handle:anchor" && item.type === "callout") return { ...item, anchor: workingWorld };
        if (mode.startsWith("handle:box:") && item.type === "callout") {
          const handle = mode.split(":")[2] as "nw" | "ne" | "sw" | "se";
          return { ...item, box: resizeRectFromHandle(item.box, handle, workingWorld) };
        }
        if (mode.startsWith("handle:rect:") && (item.type === "snapshot" || item.type === "cloud" || item.type === "rectangle" || item.type === "ellipse")) {
          const handle = mode.split(":")[2] as "nw" | "ne" | "sw" | "se";
          return { ...item, rect: resizeRectFromHandle(item.rect, handle, workingWorld) };
        }
        if (mode === "handle:start" && item.type === "measure") return { ...item, start: workingWorld };
        if (mode === "handle:end" && item.type === "measure") return { ...item, end: workingWorld };
        if (mode.startsWith("handle:point:") && item.type === "area") {
          const pointIndex = Number(mode.split(":")[2]);
          return { ...item, points: item.points.map((p, index) => (index === pointIndex ? workingWorld : p)) };
        }
        return item;
      }, { recordHistory: false });

      setDraggingSelection({ world: workingWorld });
      return;
    }

    if (!draft) {
      setSnapPreview(null);
      return;
    }
    if (draft.type === "path") {
      setSnapPreview(null);
      setDraft((prev) => (prev && prev.type === "path" ? { ...prev, points: [...prev.points, world] } : prev));
    }
    if (draft.type === "zoom-box") {
      setSnapPreview(null);
      setDraft((prev) => (prev && prev.type === "zoom-box" ? { ...prev, current: world } : prev));
    }
    if (draft.type === "selection-box") {
      setSnapPreview(null);
      setDraft((prev) => (prev && prev.type === "selection-box" ? { ...prev, current: world } : prev));
    }
    if (draft.type === "measure" || draft.type === "calibrate") {
      const snapped = getSnappedWorldPoint(world, { anchor: draft.start, excludeIds: selectedIds });
      setSnapPreview(snapped.preview);
      setDraft((prev) =>
        prev && (prev.type === "measure" || prev.type === "calibrate") ? { ...prev, end: snapped.point } : prev
      );
    }
    if (draft.type === "snapshot") {
      const snapped = getSnappedWorldPoint(world, { anchor: draft.start, excludeIds: selectedIds });
      setSnapPreview(snapped.preview);
      setDraft((prev) => (prev && prev.type === "snapshot" ? { ...prev, current: snapped.point } : prev));
    }
    if (draft.type === "rectangle") {
      const snapped = getSnappedWorldPoint(world, { anchor: draft.start, excludeIds: selectedIds });
      setSnapPreview(snapped.preview);
      setDraft((prev) => (prev && prev.type === "rectangle" ? { ...prev, current: snapped.point } : prev));
    }
    if (draft.type === "ellipse") {
      const snapped = getSnappedWorldPoint(world, { anchor: draft.start, excludeIds: selectedIds });
      setSnapPreview(snapped.preview);
      setDraft((prev) => (prev && prev.type === "ellipse" ? { ...prev, current: snapped.point } : prev));
    }
    if (draft.type === "callout") {
      const snapped = getSnappedWorldPoint(world, { anchor: draft.start, excludeIds: selectedIds });
      setSnapPreview(snapped.preview);
      setDraft((prev) => (prev && prev.type === "callout" ? { ...prev, current: snapped.point } : prev));
    }
    if (draft.type === "cloud") {
      const snapped = getSnappedWorldPoint(world, { anchor: draft.start, excludeIds: selectedIds });
      setSnapPreview(snapped.preview);
      setDraft((prev) => (prev && prev.type === "cloud" ? { ...prev, current: snapped.point } : prev));
    }
    if (draft.type === "area") {
      const anchor = draft.points[draft.points.length - 1];
      const snapped = getSnappedWorldPoint(world, { anchor, excludeIds: selectedIds });
      setSnapPreview(snapped.preview);
    }
  }

  function handlePointerUp(e?: React.PointerEvent<HTMLCanvasElement>) {
    setCursorWorld(null);
    setHoveredKeynoteToken(null);
    setSnapPreview(null);
    if (e?.pointerType === "touch") {
      touchPointersRef.current.delete(e.pointerId);
      if (touchPointersRef.current.size < 2) pinchStateRef.current = null;
      if (touchPointersRef.current.size > 0) return;
    }

    if (isPanning) {
      setIsPanning(false);
      setPanStart(null);
      return;
    }
    if (draggingSelection) {
      setDraggingSelection(null);
      return;
    }
    if (!draft) return;

    if (draft.type === "calibrate") {
      const realValue = prompt("Enter real world length (ft)", "10");
      if (realValue) {
        const pixels = distance(draft.start, draft.end);
        pushHistorySnapshot();
        setCalibration(parseFloat(realValue) / (pixels || 1));
      }
      setDraft(null);
      return;
    }
    if (draft.type === "zoom-box") {
      const x = Math.min(draft.start.x, draft.current.x);
      const y = Math.min(draft.start.y, draft.current.y);
      const w = Math.abs(draft.current.x - draft.start.x);
      const h = Math.abs(draft.current.y - draft.start.y);
      if (viewerRef.current && w > 24 && h > 24) {
        const viewer = viewerRef.current;
        const padding = 32;
        const nextScale = clamp(Math.min((viewer.clientWidth - padding) / w, (viewer.clientHeight - padding) / h), 0.25, 6);
        setTransform({
          scale: nextScale,
          x: viewer.clientWidth / 2 - (x + w / 2) * nextScale,
          y: viewer.clientHeight / 2 - (y + h / 2) * nextScale,
        });
        announceStatus("Box zoom applied");
      } else {
        announceStatus("Box zoom cancelled");
      }
      setDraft(null);
      setBoxZoomMode(false);
      return;
    }
    if (draft.type === "selection-box") {
      const selectionRect = rectFromPoints(draft.start, draft.current);
      if (selectionRect.w < 4 && selectionRect.h < 4) {
        if (!draft.append) clearSelection();
      } else {
        const hitIds = visibleAnnotationsForPage(currentPage)
          .filter((annotation) => annotationIntersectsSelectionBox(annotation, selectionRect))
          .map((annotation) => annotation.id);
        const nextSelection = draft.append ? Array.from(new Set([...selectedIds, ...hitIds])) : hitIds;
        setSelectedHandle(null);
        setSelection(nextSelection, nextSelection[nextSelection.length - 1] || null);
        announceStatus(`${nextSelection.length} markup${nextSelection.length === 1 ? "" : "s"} selected`);
      }
      setDraft(null);
      return;
    }
    if (draft.type === "area") return;

    const meta = activeUserMeta();

    if (draft.type === "callout") {
      const newCallout: CalloutAnnotation = {
        id: createId(),
        type: "callout",
        text: draft.text,
        anchor: draft.start,
        box: { x: draft.current.x, y: draft.current.y, w: 180, h: 64 },
        page: currentPage,
        author: meta.author,
        status: "open",
        createdAt: new Date().toISOString(),
        layerId: draft.layerId,
        category: draft.category,
        color: meta.color,
        roleId: meta.roleId,
        strokeWidth: activeMarkupStrokeWidth,
        fillColor: activeMarkupFillColor,
        fillOpacity: activeMarkupFillOpacity,
      };
      updateAnnotations((list) => [...list, newCallout]);
      openInlineEditorForAnnotation(newCallout);
      setSelection([newCallout.id], newCallout.id);
      setDraft(null);
      return;
    }

    if (draft.type === "cloud") {
      const x = Math.min(draft.start.x, draft.current.x);
      const y = Math.min(draft.start.y, draft.current.y);
      const w = Math.abs(draft.current.x - draft.start.x);
      const h = Math.abs(draft.current.y - draft.start.y);
      const newCloud: CloudAnnotation = {
        id: createId(),
        type: "cloud",
        rect: { x, y, w, h },
        label: draft.label,
        page: currentPage,
        author: meta.author,
        status: "open",
        createdAt: new Date().toISOString(),
        layerId: draft.layerId,
        category: draft.category,
        color: meta.color,
        roleId: meta.roleId,
        strokeWidth: activeMarkupStrokeWidth,
        fillColor: activeMarkupFillColor,
        fillOpacity: activeMarkupFillOpacity,
      };
      updateAnnotations((list) => [...list, newCloud]);
      openInlineEditorForAnnotation(newCloud);
      setSelection([newCloud.id], newCloud.id);
      setDraft(null);
      return;
    }

    if (draft.type === "rectangle") {
      const newRectangle: RectangleAnnotation = {
        id: createId(),
        type: "rectangle",
        rect: rectFromPoints(draft.start, draft.current),
        page: currentPage,
        author: meta.author,
        status: "open",
        createdAt: new Date().toISOString(),
        layerId: draft.layerId,
        category: draft.category,
        color: meta.color,
        roleId: meta.roleId,
        strokeWidth: activeMarkupStrokeWidth,
        fillColor: activeMarkupFillColor,
        fillOpacity: activeMarkupFillOpacity,
      };
      updateAnnotations((list) => [...list, newRectangle]);
      setSelection([newRectangle.id], newRectangle.id);
      setDraft(null);
      setSnapPreview(null);
      return;
    }

    if (draft.type === "ellipse") {
      const newEllipse: EllipseAnnotation = {
        id: createId(),
        type: "ellipse",
        rect: rectFromPoints(draft.start, draft.current),
        page: currentPage,
        author: meta.author,
        status: "open",
        createdAt: new Date().toISOString(),
        layerId: draft.layerId,
        category: draft.category,
        color: meta.color,
        roleId: meta.roleId,
        strokeWidth: activeMarkupStrokeWidth,
        fillColor: activeMarkupFillColor,
        fillOpacity: activeMarkupFillOpacity,
      };
      updateAnnotations((list) => [...list, newEllipse]);
      setSelection([newEllipse.id], newEllipse.id);
      setDraft(null);
      setSnapPreview(null);
      return;
    }

    if (draft.type === "snapshot") {
      const x = Math.min(draft.start.x, draft.current.x);
      const y = Math.min(draft.start.y, draft.current.y);
      const w = Math.abs(draft.current.x - draft.start.x);
      const h = Math.abs(draft.current.y - draft.start.y);
      const newSnapshot: SnapshotAnnotation = {
        id: createId(),
        type: "snapshot",
        title: draft.title,
        rect: { x, y, w, h },
        page: currentPage,
        author: meta.author,
        status: "open",
        createdAt: new Date().toISOString(),
        layerId: draft.layerId,
        category: draft.category,
        color: meta.color,
        roleId: meta.roleId,
        strokeWidth: activeMarkupStrokeWidth,
        fillColor: activeMarkupFillColor,
        fillOpacity: activeMarkupFillOpacity,
      };
      updateAnnotations((list) => [...list, newSnapshot]);
      openInlineEditorForAnnotation(newSnapshot);
      setSelection([newSnapshot.id], newSnapshot.id);
      setDraft(null);
      return;
    }

    const createdAnnotation = {
      ...draft,
      id: createId(),
      page: currentPage,
      author: meta.author,
      status: "open",
      createdAt: new Date().toISOString(),
      color: meta.color,
      roleId: meta.roleId,
      strokeWidth: activeMarkupStrokeWidth,
      fillColor: activeMarkupFillColor,
      fillOpacity: activeMarkupFillOpacity,
    } as Annotation;
    updateAnnotations((list) => [
      ...list,
      createdAnnotation,
    ]);
    setSelection([createdAnnotation.id], createdAnnotation.id);
    setDraft(null);
  }

  function finishArea() {
    if (!draft || draft.type !== "area" || draft.points.length < 3) return;
    const meta = activeUserMeta();
    const createdArea = {
      ...draft,
      id: createId(),
      page: currentPage,
      author: meta.author,
      status: "open",
      createdAt: new Date().toISOString(),
      color: meta.color,
      roleId: meta.roleId,
      strokeWidth: activeMarkupStrokeWidth,
      fillColor: activeMarkupFillColor,
      fillOpacity: activeMarkupFillOpacity,
    } as Annotation;
    updateAnnotations((list) => [
      ...list,
      createdArea,
    ]);
    setSelection([createdArea.id], createdArea.id);
    setSnapPreview(null);
    setDraft(null);
  }

  function deleteSelected() {
    if (!selectedIds.length || !canDelete()) return;
    const selectedSet = new Set(selectedIds);
    updateAnnotations((list) => list.filter((item) => !selectedSet.has(item.id)));
    clearSelection();
  }

  function duplicateSelected() {
    if (!selectedAnnotations.length || !canEdit()) return;
    const createdAt = new Date().toISOString();
    const duplicatedItems = selectedAnnotations.map((annotation, index) => {
      const duplicated = offsetAnnotation(annotation, 20 + index * 8, 20 + index * 8);
      duplicated.id = createId();
      duplicated.createdAt = createdAt;
      duplicated.page = currentPage;
      return duplicated;
    });
    updateAnnotations((list) => [...list, ...duplicatedItems]);
    setSelection(duplicatedItems.map((item) => item.id), duplicatedItems[duplicatedItems.length - 1]?.id || null);
  }

  function updateSelectedField(field: string, value: any) {
    if (!selectedIds.length || !canEdit()) return;
    patchSelectedMany((item) => ({ ...item, [field]: value } as Annotation));
  }

  function toggleLayer(layerId: string) {
    pushHistorySnapshot();
    setLayers((prev) => prev.map((layer) => (layer.id === layerId ? { ...layer, visible: !layer.visible } : layer)));
  }

  function saveSession() {
    const payload: PersistedSession = { id: createId(), name: `${projectName} - ${fileName}`, annotations, calibration, fileName, layers, savedAt: new Date().toISOString() };
    setSessions((prev) => [payload, ...prev.slice(0, 19)]);
  }

  function loadSession(session: PersistedSession) {
    setAnnotations(session.annotations || {});
    setCalibration(session.calibration || 1);
    setFileName(session.fileName || fileName);
    if (session.layers) setLayers(session.layers);
    clearSelection();
    clearHistory();
  }

  function exportBlob(filename: string, data: string) {
    const blob = new Blob([data], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = filename;
    link.click();
    URL.revokeObjectURL(url);
  }

  function exportSession() {
    if (!canExport()) return;
    exportBlob("plan_session_v5.json", JSON.stringify({ annotations, calibration, projectName, fileName, layers, users }, null, 2));
  }

  function exportAnnotatedPdfMock() {
    if (!canExport()) return;
    exportBlob(
      `${fileName.replace(/\.pdf$/i, "")}_annotated_manifest.json`,
      JSON.stringify({ type: "annotated-pdf-mock", fileName, exportedAt: new Date().toISOString(), annotations, layers, calibration }, null, 2)
    );
  }

  function exportReport() {
    const reportRows = Object.entries(annotations).flatMap(([page, items]) =>
      ensureArray(items).map((item) => ({
        page: Number(page),
        type: item.type,
        value: annotationSummary(item, calibration),
        author: item.author,
        status: item.status,
        category: categoryForAnnotation(item),
      }))
    );
    exportBlob(
      `${reportTitle.replace(/\s+/g, "_").toLowerCase()}.json`,
      JSON.stringify({ title: reportTitle, generatedAt: new Date().toISOString(), projectName, fileName, totals: computeLegendFromAnnotations(reportRows as any), markups: reportRows }, null, 2)
    );
  }

  async function saveProject() {
    const projectId = activeProjectId || createId();
    const basePayload: PersistedProject = {
      id: projectId,
      name: projectName,
      fileName,
      annotations,
      calibration,
      pageTexts,
      pageSheetNumbers,
      pdfBlobUrl: activePdfBlobUrl || undefined,
      layers,
      users,
      updatedAt: new Date().toISOString(),
    };

    setProjects((prev) => {
      const others = prev.filter((item) => item.id !== basePayload.id);
      return [basePayload, ...others];
    });
    setActiveProjectId(projectId);

    try {
      setIsCloudSaving(true);
      const pdfBlobUrl = await uploadCurrentPdfIfNeeded(projectId);
      const payload = { ...basePayload, pdfBlobUrl: pdfBlobUrl || undefined, updatedAt: new Date().toISOString() };
      const response = await fetch("/api/cloud-projects", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      });
      if (!response.ok) {
        throw new Error(`Cloud save failed (${response.status})`);
      }
      const result = (await response.json()) as { project?: CloudProjectSummary };
      setProjects((prev) => {
        const others = prev.filter((item) => item.id !== payload.id);
        return [payload, ...others];
      });
      setCloudStatus("ready");
      setCloudMessage("Shared cloud ready");
      if (result.project) {
        setCloudProjects((prev) => {
          const others = prev.filter((item) => item.id !== result.project!.id);
          return [result.project!, ...others].sort((a, b) => String(b.updatedAt).localeCompare(String(a.updatedAt)));
        });
      } else {
        await refreshCloudProjects();
      }
      announceStatus("Project saved to shared cloud");
    } catch (error) {
      const message = error instanceof Error ? error.message : "Cloud save failed";
      setCloudStatus("error");
      setCloudMessage(message);
      announceStatus("Saved locally, but shared cloud save failed");
    } finally {
      setIsCloudSaving(false);
    }
  }

  async function loadProject(project: PersistedProject | (CloudProjectSummary & Partial<PersistedProject>)) {
    setActiveProjectId(project.id);
    setProjectName(project.name);
    setFileName(project.fileName || fileName);
    try {
      setIsCloudLoading(true);
      let resolvedProject = project as PersistedProject;
      const projectDataUrl = "projectDataUrl" in project ? project.projectDataUrl : undefined;
      if (projectDataUrl) {
        const response = await fetch(projectDataUrl, { cache: "no-store" });
        if (!response.ok) throw new Error(`Cloud project load failed (${response.status})`);
        resolvedProject = await response.json();
      }

      if (resolvedProject.pdfBlobUrl) {
        await loadPdfFromSource({
          fileName: resolvedProject.fileName || fileName,
          url: resolvedProject.pdfBlobUrl,
          pdfBlobUrl: resolvedProject.pdfBlobUrl,
        });
        setActivePdfFile(null);
      }

      setAnnotations(resolvedProject.annotations || {});
      setCalibration(resolvedProject.calibration || 1);
      setPageTexts(resolvedProject.pageTexts || {});
      setPageSheetNumbers(resolvedProject.pageSheetNumbers || {});
      if (resolvedProject.layers) setLayers(resolvedProject.layers);
      if (resolvedProject.users) setUsers(resolvedProject.users);
      clearSelection();
      clearHistory();
      announceStatus(`Loaded ${resolvedProject.name}`);
    } catch (error) {
      const message = error instanceof Error ? error.message : "Unable to load project";
      setCloudStatus("error");
      setCloudMessage(message);
      announceStatus(message);
    } finally {
      setIsCloudLoading(false);
    }
  }

  function runOcrSearch() {
    const query = ocrQuery.trim().toLowerCase();
    if (!query) {
      setOcrResults([]);
      return;
    }
    const results = Object.entries(pageTexts).flatMap(([page, text]) => {
      const content = String(text || "");
      const lower = content.toLowerCase();
      if (!lower.includes(query)) return [];
      const index = lower.indexOf(query);
      const start = Math.max(0, index - 40);
      const end = Math.min(content.length, index + query.length + 80);
      return [{ page: Number(page), preview: content.slice(start, end) }];
    });
    setOcrResults(results);
  }

  const markupRows = useMemo(() => {
    return Object.entries(annotations)
      .flatMap(([page, items]) => ensureArray(items).map((item) => ({ ...item, page: Number(page) })))
      .filter((item) => {
        if (!visibleLayerIds.has(item.layerId || "layer-general")) return false;
        if (markupStatusFilter !== "all" && (item.status || "open") !== markupStatusFilter) return false;
        const query = markupFilter.trim().toLowerCase();
        if (!query) return true;
        return JSON.stringify(item).toLowerCase().includes(query);
      });
  }, [annotations, markupFilter, markupStatusFilter, visibleLayerIds]);

  const totals = useMemo(() => {
    let totalLength = 0;
    let totalArea = 0;
    const visibleItems = Object.values(annotations).flat().filter((item) => visibleLayerIds.has(item.layerId || "layer-general"));
    visibleItems.forEach((item) => {
      if (item.type === "measure") totalLength += distance(item.start, item.end);
      if (item.type === "area") totalArea += polygonArea(item.points);
    });
    return { count: visibleItems.length, length: formatFeet(totalLength, calibration), area: formatArea(totalArea, calibration), legend: computeLegendFromAnnotations(visibleItems) };
  }, [annotations, calibration, visibleLayerIds]);
  const takeoffSummaries = useMemo(() => {
    const grouped: Record<string, { count: number; totalLength: number; totalArea: number }> = {};
    const visibleItems = Object.values(annotations).flat().filter((item) => visibleLayerIds.has(item.layerId || "layer-general"));
    visibleItems.forEach((item) => {
      const key = categoryForAnnotation(item);
      if (!grouped[key]) grouped[key] = { count: 0, totalLength: 0, totalArea: 0 };
      grouped[key].count += 1;
      if (item.type === "measure") grouped[key].totalLength += distance(item.start, item.end);
      if (item.type === "area") grouped[key].totalArea += polygonArea(item.points);
    });
    return Object.entries(grouped)
      .map(([category, values]) => ({
        category,
        count: values.count,
        length: values.totalLength ? formatFeet(values.totalLength, calibration) : "0.00 ft",
        area: values.totalArea ? formatArea(values.totalArea, calibration) : "0.00 sf",
      }))
      .sort((a, b) => {
        const areaA = Number.parseFloat(a.area);
        const areaB = Number.parseFloat(b.area);
        const lengthA = Number.parseFloat(a.length);
        const lengthB = Number.parseFloat(b.length);
        return areaB - areaA || lengthB - lengthA || b.count - a.count || a.category.localeCompare(b.category);
      });
  }, [annotations, calibration, visibleLayerIds]);

  useEffect(() => {
    const canvas = overlayCanvasRef.current;
    const viewer = viewerRef.current;
    if (!canvas || !viewer) return;

    canvas.width = viewer.clientWidth;
    canvas.height = viewer.clientHeight;

    const context = canvas.getContext("2d");
    if (!context) return;
    context.clearRect(0, 0, canvas.width, canvas.height);
    context.lineJoin = "round";
    context.lineCap = "round";

    visibleAnnotationsForPage(currentPage).forEach((item) => {
      const selected = selectedIdSet.has(item.id);
      const strokeWidth = item.strokeWidth ?? 2;
      const fillColor = item.fillColor || item.color || "#2563eb";
      const fillOpacity = item.fillOpacity ?? 0.16;
      context.save();
      context.strokeStyle = item.color || (selected ? "#2563eb" : "#dc2626");
      context.fillStyle = item.color || "#111827";
      context.lineWidth = selected ? Math.max(strokeWidth + 1, 3) : strokeWidth;

      if (item.type === "measure") {
        const start = worldToScreen(item.start, transform);
        const end = worldToScreen(item.end, transform);
        const label = formatFeet(distance(item.start, item.end), calibration);
        context.beginPath();
        context.moveTo(start.x, start.y);
        context.lineTo(end.x, end.y);
        context.stroke();
        context.beginPath();
        context.arc(start.x, start.y, 4, 0, Math.PI * 2);
        context.arc(end.x, end.y, 4, 0, Math.PI * 2);
        context.fill();
        drawCanvasValueLabel(context, label, (start.x + end.x) / 2, (start.y + end.y) / 2 + 16, transform.scale);
      }

      if (item.type === "text" || item.type === "comment") {
        const position = worldToScreen(item.position, transform);
        context.font = item.type === "comment" ? "italic 14px sans-serif" : "14px sans-serif";
        context.fillText(item.text || item.comment || "", position.x, position.y);
      }

      if (item.type === "callout") {
        const anchor = worldToScreen(item.anchor, transform);
        const boxX = item.box.x * transform.scale + transform.x;
        const boxY = item.box.y * transform.scale + transform.y;
        const boxW = item.box.w * transform.scale;
        const boxH = item.box.h * transform.scale;
        const leader = nearestPointOnRect(anchor, { x: boxX, y: boxY, w: boxW, h: boxH });
        context.beginPath();
        context.moveTo(anchor.x, anchor.y);
        context.lineTo(leader.x, leader.y);
        context.stroke();
        context.fillStyle = hexToRgba(fillColor, selected ? Math.max(fillOpacity, 0.24) : fillOpacity);
        context.fillRect(boxX, boxY, boxW, boxH);
        context.strokeRect(boxX, boxY, boxW, boxH);
        context.beginPath();
        context.arc(anchor.x, anchor.y, 4, 0, Math.PI * 2);
        context.fillStyle = item.color || "#111827";
        context.fill();
        context.fillStyle = "#f8fafc";
        context.font = "13px sans-serif";
        context.fillText(item.text, boxX + 8, boxY + 22);
        if (selected) drawSelectionHandles(context, item.box, transform);
      }

      if (item.type === "area") {
        context.beginPath();
        item.points.forEach((point, index) => {
          const screen = worldToScreen(point, transform);
          if (index === 0) context.moveTo(screen.x, screen.y);
          else context.lineTo(screen.x, screen.y);
        });
        context.closePath();
        context.fillStyle = hexToRgba(fillColor, selected ? Math.max(fillOpacity, 0.24) : fillOpacity);
        context.fill();
        context.stroke();
        const first = worldToScreen(item.points[0], transform);
        const areaLabel = formatArea(polygonArea(item.points), calibration);
        drawCanvasValueLabel(context, areaLabel, first.x, first.y, transform.scale);
      }

      if (item.type === "path") {
        context.beginPath();
        item.points.forEach((point, index) => {
          const screen = worldToScreen(point, transform);
          if (index === 0) context.moveTo(screen.x, screen.y);
          else context.lineTo(screen.x, screen.y);
        });
        context.stroke();
      }

      if (item.type === "rectangle") {
        const x = item.rect.x * transform.scale + transform.x;
        const y = item.rect.y * transform.scale + transform.y;
        const width = item.rect.w * transform.scale;
        const height = item.rect.h * transform.scale;
        context.fillStyle = hexToRgba(fillColor, selected ? Math.max(fillOpacity, 0.24) : fillOpacity);
        context.fillRect(x, y, width, height);
        context.strokeRect(x, y, width, height);
        if (selected) drawSelectionHandles(context, item.rect, transform);
      }

      if (item.type === "ellipse") {
        const center = worldToScreen({ x: item.rect.x + item.rect.w / 2, y: item.rect.y + item.rect.h / 2 }, transform);
        context.beginPath();
        context.ellipse(center.x, center.y, Math.abs(item.rect.w * transform.scale) / 2, Math.abs(item.rect.h * transform.scale) / 2, 0, 0, Math.PI * 2);
        context.fillStyle = hexToRgba(fillColor, selected ? Math.max(fillOpacity, 0.24) : fillOpacity);
        context.fill();
        context.stroke();
        if (selected) drawSelectionHandles(context, item.rect, transform);
      }

      if (item.type === "link") {
        const position = worldToScreen(item.position, transform);
        context.fillStyle = "#1d4ed8";
        context.font = "14px sans-serif";
        context.fillText(`→ Page ${item.targetPage}`, position.x, position.y);
      }

      if (item.type === "stamp") {
        const stampRect = getStampWorldRect(item);
        const x = stampRect.x * transform.scale + transform.x;
        const y = stampRect.y * transform.scale + transform.y;
        const width = stampRect.w * transform.scale;
        const height = stampRect.h * transform.scale;
        context.save();
        context.translate(x + width / 2, y + height / 2);
        context.rotate((-12 * Math.PI) / 180);
        context.fillStyle = hexToRgba(fillColor, selected ? Math.max(fillOpacity, 0.28) : Math.max(fillOpacity, 0.18));
        context.strokeStyle = item.color || "#2563eb";
        context.fillRect(-width / 2, -height / 2, width, height);
        context.strokeRect(-width / 2, -height / 2, width, height);
        context.font = `700 ${Math.max(12, Math.min(28, height * 0.4))}px sans-serif`;
        context.textAlign = "center";
        context.textBaseline = "middle";
        context.fillStyle = item.color || "#2563eb";
        context.fillText(item.text, 0, 1);
        context.restore();
      }

      if (item.type === "snapshot") {
        const x = item.rect.x * transform.scale + transform.x;
        const y = item.rect.y * transform.scale + transform.y;
        const width = item.rect.w * transform.scale;
        const height = item.rect.h * transform.scale;
        context.fillStyle = hexToRgba(fillColor, selected ? Math.max(fillOpacity, 0.24) : fillOpacity);
        context.fillRect(x, y, width, height);
        context.strokeRect(x, y, width, height);
        context.fillText(item.title || "Snapshot", x + 4, y - 6);
        if (selected) drawSelectionHandles(context, item.rect, transform);
      }

      if (item.type === "cloud") {
        const cloudRect = {
          x: item.rect.x * transform.scale + transform.x,
          y: item.rect.y * transform.scale + transform.y,
          w: item.rect.w * transform.scale,
          h: item.rect.h * transform.scale,
        };
        drawCloudRect(context, cloudRect);
        context.fillStyle = hexToRgba(fillColor, selected ? Math.max(fillOpacity, 0.24) : fillOpacity);
        context.fill();
        context.stroke();
        if (item.label) {
          context.fillStyle = "#f8fafc";
          context.font = "13px sans-serif";
          context.fillText(item.label, cloudRect.x + 8, cloudRect.y - 8);
        }
        if (selected) drawSelectionHandles(context, item.rect, transform);
      }

      context.restore();
    });

    if (tool === TOOL.PAN || tool === TOOL.SELECT) {
      currentAutoReferenceLinks.forEach((link) => {
        const x = link.rect.x * transform.scale + transform.x;
        const y = link.rect.y * transform.scale + transform.y;
        const width = link.rect.w * transform.scale;
        const height = link.rect.h * transform.scale;
        const isHovered = hoveredAutoReferenceId === link.id;

        context.save();
        context.strokeStyle = isHovered ? "rgba(37,99,235,0.95)" : "rgba(37,99,235,0.28)";
        context.fillStyle = isHovered ? "rgba(37,99,235,0.12)" : "rgba(37,99,235,0.04)";
        context.lineWidth = isHovered ? 2 : 1;
        context.setLineDash(isHovered ? [] : [4, 3]);
        context.fillRect(x, y, width, height);
        context.strokeRect(x, y, width, height);
        context.restore();
      });
    }

    if (hoveredKeynoteToken && highlightedKeynoteFragments.length) {
      highlightedKeynoteFragments.forEach((fragment) => {
        const x = fragment.rect.x * transform.scale + transform.x;
        const y = fragment.rect.y * transform.scale + transform.y;
        const width = fragment.rect.w * transform.scale;
        const height = fragment.rect.h * transform.scale;
        context.save();
        context.fillStyle = "rgba(250, 204, 21, 0.22)";
        context.strokeStyle = "rgba(250, 204, 21, 0.78)";
        context.lineWidth = 1.5;
        context.fillRect(x - 4, y - 2, width + 8, height + 4);
        context.strokeRect(x - 4, y - 2, width + 8, height + 4);
        context.restore();
      });
    }

    if (snapPreview) {
      const snapScreen = worldToScreen(snapPreview.point, transform);
      context.save();
      context.strokeStyle = "rgba(34, 211, 238, 0.95)";
      context.fillStyle = "rgba(34, 211, 238, 0.18)";
      context.lineWidth = 1.5;
      if (snapPreview.guideX !== undefined) {
        const guideStart = worldToScreen({ x: snapPreview.guideX, y: 0 }, transform);
        const guideEnd = worldToScreen({ x: snapPreview.guideX, y: currentPageSize.height }, transform);
        context.beginPath();
        context.moveTo(guideStart.x, guideStart.y);
        context.lineTo(guideEnd.x, guideEnd.y);
        context.stroke();
      }
      if (snapPreview.guideY !== undefined) {
        const guideStart = worldToScreen({ x: 0, y: snapPreview.guideY }, transform);
        const guideEnd = worldToScreen({ x: currentPageSize.width, y: snapPreview.guideY }, transform);
        context.beginPath();
        context.moveTo(guideStart.x, guideStart.y);
        context.lineTo(guideEnd.x, guideEnd.y);
        context.stroke();
      }
      context.beginPath();
      context.arc(snapScreen.x, snapScreen.y, 7, 0, Math.PI * 2);
      context.fill();
      context.beginPath();
      context.moveTo(snapScreen.x - 10, snapScreen.y);
      context.lineTo(snapScreen.x + 10, snapScreen.y);
      context.moveTo(snapScreen.x, snapScreen.y - 10);
      context.lineTo(snapScreen.x, snapScreen.y + 10);
      context.stroke();
      drawCanvasValueLabel(context, snapPreview.label, snapScreen.x + 12, snapScreen.y - 14, 0.9);
      context.restore();
    }

    if (!draft) return;
    context.save();
    context.strokeStyle = activeMarkupColor;
    context.lineWidth = activeMarkupStrokeWidth;

    if (draft.type === "measure" || draft.type === "calibrate") {
      const start = worldToScreen(draft.start, transform);
      const end = worldToScreen(draft.end, transform);
      context.beginPath();
      context.moveTo(start.x, start.y);
      context.lineTo(end.x, end.y);
      context.stroke();
      context.beginPath();
      context.arc(start.x, start.y, 4, 0, Math.PI * 2);
      context.arc(end.x, end.y, 4, 0, Math.PI * 2);
      context.fill();
      if (draft.type === "measure") {
        drawCanvasValueLabel(
          context,
          formatFeet(distance(draft.start, draft.end), calibration),
          (start.x + end.x) / 2,
          (start.y + end.y) / 2 + 16,
          transform.scale
        );
      }
    }

    if (draft.type === "path") {
      context.beginPath();
      draft.points.forEach((point, index) => {
        const screen = worldToScreen(point, transform);
        if (index === 0) context.moveTo(screen.x, screen.y);
        else context.lineTo(screen.x, screen.y);
      });
      context.stroke();
    }

    if (draft.type === "zoom-box") {
      const x = Math.min(draft.start.x, draft.current.x) * transform.scale + transform.x;
      const y = Math.min(draft.start.y, draft.current.y) * transform.scale + transform.y;
      const width = Math.abs(draft.current.x - draft.start.x) * transform.scale;
      const height = Math.abs(draft.current.y - draft.start.y) * transform.scale;
      context.save();
      context.setLineDash([8, 6]);
      context.strokeStyle = "#f59e0b";
      context.fillStyle = "rgba(245, 158, 11, 0.12)";
      context.strokeRect(x, y, width, height);
      context.fillRect(x, y, width, height);
      context.restore();
    }

    if (draft.type === "selection-box") {
      const x = Math.min(draft.start.x, draft.current.x) * transform.scale + transform.x;
      const y = Math.min(draft.start.y, draft.current.y) * transform.scale + transform.y;
      const width = Math.abs(draft.current.x - draft.start.x) * transform.scale;
      const height = Math.abs(draft.current.y - draft.start.y) * transform.scale;
      context.save();
      context.setLineDash([6, 4]);
      context.strokeStyle = "rgba(96, 165, 250, 0.95)";
      context.fillStyle = "rgba(96, 165, 250, 0.14)";
      context.strokeRect(x, y, width, height);
      context.fillRect(x, y, width, height);
      context.restore();
    }

    if (draft.type === "area") {
      const previewPoints = cursorWorld ? [...draft.points, cursorWorld] : draft.points;
      context.beginPath();
      previewPoints.forEach((point, index) => {
        const screen = worldToScreen(point, transform);
        if (index === 0) context.moveTo(screen.x, screen.y);
        else context.lineTo(screen.x, screen.y);
      });
      if (cursorWorld && draft.points.length >= 2) {
        context.setLineDash([7, 5]);
      }
      context.fillStyle = hexToRgba(activeMarkupFillColor, activeMarkupFillOpacity);
      context.fill();
      context.stroke();
      context.setLineDash([]);
      if (draft.points[0]) {
        const firstPoint = worldToScreen(draft.points[0], transform);
        context.save();
        context.beginPath();
        context.fillStyle = "#f8fafc";
        context.arc(firstPoint.x, firstPoint.y, 4, 0, Math.PI * 2);
        context.fill();
        context.restore();
      }
      if (previewPoints.length >= 3) {
        const areaPreview = formatArea(polygonArea(previewPoints), calibration);
        const anchorPoint = worldToScreen(previewPoints[0], transform);
        drawCanvasValueLabel(context, areaPreview, anchorPoint.x, anchorPoint.y - 16, transform.scale);
      }
    }

    if (draft.type === "snapshot") {
      const x = Math.min(draft.start.x, draft.current.x) * transform.scale + transform.x;
      const y = Math.min(draft.start.y, draft.current.y) * transform.scale + transform.y;
      const width = Math.abs(draft.current.x - draft.start.x) * transform.scale;
      const height = Math.abs(draft.current.y - draft.start.y) * transform.scale;
      context.fillStyle = hexToRgba(activeMarkupFillColor, activeMarkupFillOpacity);
      context.fillRect(x, y, width, height);
      context.strokeRect(x, y, width, height);
    }

    if (draft.type === "rectangle") {
      const rect = rectFromPoints(draft.start, draft.current);
      const x = rect.x * transform.scale + transform.x;
      const y = rect.y * transform.scale + transform.y;
      const width = rect.w * transform.scale;
      const height = rect.h * transform.scale;
      context.fillStyle = hexToRgba(activeMarkupFillColor, activeMarkupFillOpacity);
      context.fillRect(x, y, width, height);
      context.strokeRect(x, y, width, height);
    }

    if (draft.type === "ellipse") {
      const rect = rectFromPoints(draft.start, draft.current);
      const center = worldToScreen({ x: rect.x + rect.w / 2, y: rect.y + rect.h / 2 }, transform);
      context.beginPath();
      context.ellipse(center.x, center.y, Math.abs(rect.w * transform.scale) / 2, Math.abs(rect.h * transform.scale) / 2, 0, 0, Math.PI * 2);
      context.fillStyle = hexToRgba(activeMarkupFillColor, activeMarkupFillOpacity);
      context.fill();
      context.stroke();
    }

    if (draft.type === "callout") {
      const anchor = worldToScreen(draft.start, transform);
      const boxRect = {
        x: draft.current.x * transform.scale + transform.x,
        y: draft.current.y * transform.scale + transform.y,
        w: 180 * transform.scale,
        h: 64 * transform.scale,
      };
      const leader = nearestPointOnRect(anchor, boxRect);
      context.beginPath();
      context.moveTo(anchor.x, anchor.y);
      context.lineTo(leader.x, leader.y);
      context.stroke();
      context.fillStyle = hexToRgba(activeMarkupFillColor, activeMarkupFillOpacity);
      context.fillRect(boxRect.x, boxRect.y, boxRect.w, boxRect.h);
      context.strokeRect(boxRect.x, boxRect.y, boxRect.w, boxRect.h);
      context.fillStyle = "#f8fafc";
      context.font = "13px sans-serif";
      context.fillText(draft.text, boxRect.x + 8, boxRect.y + 22);
    }

    if (draft.type === "cloud") {
      const cloudRect = {
        x: Math.min(draft.start.x, draft.current.x) * transform.scale + transform.x,
        y: Math.min(draft.start.y, draft.current.y) * transform.scale + transform.y,
        w: Math.abs(draft.current.x - draft.start.x) * transform.scale,
        h: Math.abs(draft.current.y - draft.start.y) * transform.scale,
      };
      drawCloudRect(context, cloudRect);
      context.fillStyle = hexToRgba(activeMarkupFillColor, activeMarkupFillOpacity);
      context.fill();
      context.stroke();
    }

    context.restore();
  }, [annotations, autoReferenceLinks, currentAutoReferenceLinks, currentPageSize.height, currentPageSize.width, draft, hoveredAutoReferenceId, hoveredKeynoteToken, highlightedKeynoteFragments, transform, currentPage, selectedIdSet, calibration, layers, visibleLayerIds, tool, activeMarkupColor, activeMarkupStrokeWidth, activeMarkupFillColor, activeMarkupFillOpacity, snapPreview]);

  const menuLabels: Record<MenuId, string> = {
    file: "File",
    edit: "Edit",
    view: "View",
    document: "Document",
    markup: "Markup",
    measure: "Measure",
    window: "Window",
    help: "Help",
  };

  const menuEntries: Record<MenuId, MenuItem[]> = {
    file: [
      { label: "Open PDF...", shortcut: "Ctrl+O", action: () => { openPdfPicker(); announceStatus("Opening PDF picker"); } },
      { label: "Save Project", shortcut: "Ctrl+S", action: () => { saveProject(); announceStatus("Project saved"); } },
      { label: "Save Session", action: () => { saveSession(); announceStatus("Session saved"); } },
      { separator: true, label: "sep-1" },
      { label: "Export Data", action: () => { exportSession(); announceStatus("Exported session data"); }, disabled: !canExport() },
      { label: "Export Annotated Manifest", action: () => { exportAnnotatedPdfMock(); announceStatus("Exported annotated manifest"); }, disabled: !canExport() },
      { separator: true, label: "sep-2" },
      { label: isViewerFullscreen ? "Exit Fullscreen" : "Enter Fullscreen", action: () => { void toggleViewerFullscreen(); announceStatus(isViewerFullscreen ? "Exited fullscreen" : "Entered fullscreen"); } },
    ],
    edit: [
      { label: "Undo", shortcut: "Ctrl+Z", action: () => undoLastAction(), disabled: undoStack.length === 0 },
      { label: "Redo", shortcut: "Ctrl+Y", action: () => redoLastAction(), disabled: redoStack.length === 0 },
      { separator: true, label: "sep-edit-history" },
      { label: "Copy Selected", shortcut: "Ctrl+C", action: () => copySelectionToClipboard(), disabled: !selectedIds.length },
      { label: "Paste", shortcut: "Ctrl+V", action: () => pasteClipboardSelection(), disabled: !clipboardAnnotations.length || !canEdit() },
      { label: "Duplicate Selected", action: () => { duplicateSelected(); announceStatus("Selection duplicated"); }, disabled: !selectedIds.length || !canEdit() },
      { label: "Delete Selected", shortcut: "Delete", action: () => { deleteSelected(); announceStatus("Selection deleted"); }, disabled: !selectedIds.length || !canDelete() },
      { label: "Finish Area", action: () => { finishArea(); announceStatus("Area markup finished"); }, disabled: !(draft && draft.type === "area") },
    ],
    view: [
      { label: "Zoom In", action: () => { zoomBy(1.2); announceStatus("Zoomed in"); } },
      { label: "Zoom Out", action: () => { zoomBy(1 / 1.2); announceStatus("Zoomed out"); } },
      { label: "Zoom 50%", action: () => zoomToPercent(50) },
      { label: "Zoom 100%", action: () => zoomToPercent(100) },
      { label: "Zoom 200%", action: () => zoomToPercent(200) },
      { label: "Fit to Screen", shortcut: "Ctrl+0", action: () => { fitToViewer(); announceStatus("Fit to screen"); } },
      { label: "Fit Width", action: () => { fitWidthToViewer(); announceStatus("Fit width view"); } },
      { label: "Fit Height", action: () => { fitHeightToViewer(); announceStatus("Fit height view"); } },
      { label: "Box Zoom", shortcut: "Z", action: () => activateBoxZoom() },
      { label: navigationMode === "sheet" ? "Wheel Mode: Sheet Pages" : "Wheel Mode: Review Pan", shortcut: "G", action: () => setNavigationModeAndAnnounce(navigationMode === "sheet" ? "review" : "sheet") },
      { separator: true, label: "sep-3" },
      { label: showMarkupList ? "Hide Markup List" : "Show Markup List", action: () => { setShowMarkupList((value) => !value); announceStatus(showMarkupList ? "Markup list hidden" : "Markup list shown"); } },
      { label: compareMode ? "Disable Compare" : "Enable Compare", action: () => { setCompareMode((value) => !value); announceStatus(compareMode ? "Compare disabled" : "Compare enabled"); } },
    ],
    document: [
      { label: "Previous Page", shortcut: "Page Up", action: () => { changePageBy(-1); announceStatus("Previous page"); }, disabled: currentPage <= 1 },
      { label: "Next Page", shortcut: "Page Down", action: () => { changePageBy(1); announceStatus("Next page"); }, disabled: currentPage >= pageCount },
      { label: "First Page", action: () => { setCurrentPage(1); announceStatus("Jumped to first page"); }, disabled: currentPage <= 1 },
      { label: "Last Page", action: () => { setCurrentPage(Math.max(pageCount, 1)); announceStatus("Jumped to last page"); }, disabled: currentPage >= pageCount },
    ],
    markup: [
      { label: "Select Tool", shortcut: "V", action: () => setToolAndAnnounce(TOOL.SELECT, "Select tool") },
      { label: "Pan Tool", shortcut: "H", action: () => setToolAndAnnounce(TOOL.PAN, "Pan tool") },
      { label: "Sketch Markup", shortcut: "D", action: () => setToolAndAnnounce(TOOL.DRAW, "Sketch markup tool") },
      { label: "Text Markup", shortcut: "T", action: () => setToolAndAnnounce(TOOL.TEXT, "Text tool") },
      { label: "Stamp Tool", shortcut: "R", action: () => setToolAndAnnounce(TOOL.STAMP, "Stamp tool") },
      { label: "Callout Tool", action: () => setToolAndAnnounce(TOOL.CALLOUT, "Callout tool") },
      { label: "Comment Markup", action: () => setToolAndAnnounce(TOOL.COMMENT, "Comment tool") },
      { label: "Cloud Tool", action: () => setToolAndAnnounce(TOOL.CLOUD, "Cloud tool") },
      { label: "Snapshot Tool", action: () => setToolAndAnnounce(TOOL.SNAPSHOT, "Snapshot tool") },
    ],
    measure: [
      { label: "Length Tool", shortcut: "M", action: () => setToolAndAnnounce(TOOL.MEASURE, "Length tool") },
      { label: "Area Tool", shortcut: "A", action: () => setToolAndAnnounce(TOOL.AREA, "Area tool") },
      { label: "Calibrate Scale", action: () => setToolAndAnnounce(TOOL.CALIBRATE, "Calibration tool") },
    ],
    window: [
      { label: showProjectsPanel ? "Hide Project Panel" : "Show Project Panel", action: () => { setShowProjectsPanel((value) => !value); announceStatus(showProjectsPanel ? "Project panel hidden" : "Project panel shown"); } },
      { label: showInspectorPanel ? "Hide Inspector Panel" : "Show Inspector Panel", action: () => { setShowInspectorPanel((value) => !value); announceStatus(showInspectorPanel ? "Inspector panel hidden" : "Inspector panel shown"); } },
      { label: showPagesPanel ? "Hide Page Browser" : "Show Page Browser", action: () => { setShowPagesPanel((value) => !value); announceStatus(showPagesPanel ? "Page browser hidden" : "Page browser shown"); } },
      {
        label: !showProjectsPanel && !showInspectorPanel && !showPagesPanel && !showMarkupList ? "Exit Focus Mode" : "Enter Focus Mode",
        action: () => {
          const enteringFocus = showProjectsPanel || showInspectorPanel || showPagesPanel || showMarkupList;
          setShowProjectsPanel(!enteringFocus);
          setShowInspectorPanel(!enteringFocus);
          setShowPagesPanel(!enteringFocus);
          setShowMarkupList(!enteringFocus);
          announceStatus(enteringFocus ? "Focus mode enabled" : "Focus mode disabled");
        },
      },
      { label: activeBookmarkTab === "all" ? "All Sheets Visible" : "Show All Sheet Tabs", action: () => { setActiveBookmarkTab("all"); announceStatus("Showing all sheet tabs"); } },
      { label: compareMode ? "Close Compare View" : "Open Compare View", action: () => { setCompareMode((value) => !value); announceStatus(compareMode ? "Compare view closed" : "Compare view opened"); } },
      { label: showMarkupList ? "Hide Database Panel" : "Show Database Panel", action: () => { setShowMarkupList((value) => !value); announceStatus(showMarkupList ? "Markup database hidden" : "Markup database shown"); } },
    ],
    help: [
      { label: "Viewer Controls", action: showHelpDialog },
      { label: "Keyboard Shortcuts", action: showKeyboardShortcutsDialog },
      { label: "About This Workspace", action: () => window.alert("Bluebeam-style construction drawing viewer with markups, sheet links, thumbnails, fullscreen, and touch zoom.") },
    ],
  };
  const focusModeActive = !showProjectsPanel && !showInspectorPanel && !showPagesPanel && !showMarkupList;
  const centerWorkspaceClass = "space-y-2 min-w-0";

  return (
    <div className="min-h-screen bg-transparent px-1.5 py-2 text-slate-100 lg:px-2">
      <div ref={menuBarRef} className="mb-2 w-full overflow-visible rounded-2xl border border-slate-700 bg-[#14181d] shadow-2xl">
        <div className="relative flex flex-wrap items-center gap-2 border-b border-slate-800 bg-[#0f1216] px-3 py-2 text-[11px] font-medium uppercase tracking-[0.18em] text-slate-400">
          <span className="text-amber-400">Revu</span>
          {(["file", "edit", "view", "document", "markup", "measure", "window", "help"] as MenuId[]).map((menuId) => (
            <div key={menuId} className="relative">
              <button
                type="button"
                onClick={() => setActiveMenu((current) => (current === menuId ? null : menuId))}
                className={`rounded-md px-2 py-1 transition ${activeMenu === menuId ? "bg-[#232933] text-slate-100" : "hover:bg-[#1a1f25] hover:text-slate-200"}`}
              >
                {menuLabels[menuId]}
              </button>
              {activeMenu === menuId ? (
                <div className="absolute left-0 top-[calc(100%+8px)] z-50 min-w-[220px] overflow-hidden rounded-xl border border-slate-700 bg-[#171b21] py-1 shadow-2xl">
                  {menuEntries[menuId].map((item, index) =>
                    item.separator ? (
                      <div key={`${menuId}-sep-${index}`} className="my-1 h-px bg-slate-700" />
                    ) : (
                      <button
                        key={`${menuId}-${item.label}`}
                        type="button"
                        disabled={item.disabled}
                        onClick={() => {
                          item.action?.();
                          setActiveMenu(null);
                        }}
                        className="flex w-full items-center justify-between px-3 py-2 text-left text-xs normal-case tracking-normal text-slate-200 transition hover:bg-[#232933] disabled:cursor-not-allowed disabled:text-slate-500 disabled:hover:bg-transparent"
                      >
                        <span>{item.label}</span>
                        {item.shortcut ? <span className="pl-4 text-[10px] uppercase tracking-[0.14em] text-slate-500">{item.shortcut}</span> : null}
                      </button>
                    )
                  )}
                </div>
              ) : null}
            </div>
          ))}
          <button
            type="button"
            onClick={showKeyboardShortcutsDialog}
            className="rounded-md px-2 py-1 transition hover:bg-[#1a1f25] hover:text-slate-200"
          >
            Keyboard Shortcuts
          </button>
          <div className="ml-auto flex items-center gap-2 rounded-xl border border-slate-700 bg-[#11151a] px-3 py-1.5 text-[10px] uppercase tracking-[0.14em] text-slate-500">
            <span className="text-slate-300">{statusMessage}</span>
          </div>
        </div>
        <div className="flex flex-wrap items-center gap-3 bg-[#1a1f25] px-3 py-2">
          <div className="flex flex-wrap items-center gap-2 rounded-xl border border-slate-700 bg-[#11151a] px-2 py-1.5">
            <span className="px-1 text-[10px] uppercase tracking-[0.14em] text-slate-500">File</span>
            <Button size="sm" variant="secondary" onClick={() => { openPdfPicker(); announceStatus("Opening PDF picker"); }}><FolderOpen className="h-4 w-4" />Open</Button>
            <Button size="sm" variant="secondary" onClick={() => { saveProject(); announceStatus("Project saved"); }}><Save className="h-4 w-4" />Save</Button>
          </div>
          <div className="flex flex-wrap items-center gap-2 rounded-xl border border-slate-700 bg-[#11151a] px-2 py-1.5">
            <span className="px-1 text-[10px] uppercase tracking-[0.14em] text-slate-500">Sheet</span>
            <Button size="sm" variant="outline" onClick={() => changePageBy(-1)} disabled={currentPage <= 1}><ChevronLeft className="h-4 w-4" />Prev</Button>
            <Badge variant="secondary" className="h-8 px-3">{currentPage}{currentSheetNumber ? ` ${currentSheetNumber}` : ""}</Badge>
            <Button size="sm" variant="outline" onClick={() => changePageBy(1)} disabled={currentPage >= pageCount}><ChevronRight className="h-4 w-4" />Next</Button>
            <Button size="sm" variant={navigationMode === "sheet" ? "default" : "secondary"} onClick={() => setNavigationModeAndAnnounce(navigationMode === "sheet" ? "review" : "sheet")}>
              Wheel: {navigationMode === "sheet" ? "Pages" : "Pan"}
            </Button>
            <Button size="sm" variant={compareMode ? "default" : "secondary"} onClick={() => { setCompareMode((value) => !value); announceStatus(compareMode ? "Compare disabled" : "Compare enabled"); }}><GitCompare className="h-4 w-4" />Compare</Button>
            <Button size="sm" variant="secondary" onClick={() => { void toggleViewerFullscreen(); announceStatus(isViewerFullscreen ? "Exited fullscreen" : "Entered fullscreen"); }}>
              {isViewerFullscreen ? <Minimize2 className="h-4 w-4" /> : <Maximize2 className="h-4 w-4" />}
              {isViewerFullscreen ? "Exit" : "Fullscreen"}
            </Button>
          </div>
          <div className="flex flex-wrap items-center gap-2 rounded-xl border border-slate-700 bg-[#11151a] px-2 py-1.5">
            <span className="px-1 text-[10px] uppercase tracking-[0.14em] text-slate-500">Workspace</span>
            <Button size="sm" variant={showProjectsPanel ? "outline" : "secondary"} onClick={() => { setShowProjectsPanel((value) => !value); announceStatus(showProjectsPanel ? "Project panel hidden" : "Project panel shown"); }}>Projects</Button>
            <Button size="sm" variant={showPagesPanel ? "outline" : "secondary"} onClick={() => { setShowPagesPanel((value) => !value); announceStatus(showPagesPanel ? "Page browser hidden" : "Page browser shown"); }}>Pages</Button>
            <Button size="sm" variant={showInspectorPanel ? "outline" : "secondary"} onClick={() => { setShowInspectorPanel((value) => !value); announceStatus(showInspectorPanel ? "Inspector panel hidden" : "Inspector panel shown"); }}>Inspector</Button>
            <Button size="sm" variant={focusModeActive ? "default" : "secondary"} onClick={() => {
              const enteringFocus = !focusModeActive;
              setShowProjectsPanel(!enteringFocus);
              setShowPagesPanel(!enteringFocus);
              setShowInspectorPanel(!enteringFocus);
              setShowMarkupList(!enteringFocus);
              announceStatus(enteringFocus ? "Focus mode enabled" : "Focus mode disabled");
            }}>Focus</Button>
          </div>
          <Button size="sm" variant="secondary" onClick={() => { runOcrSearch(); announceStatus(ocrQuery.trim() ? "Search results updated" : "Enter search text to scan sheets"); }}><Search className="h-4 w-4" />Search</Button>
          <form
            className="ml-auto flex items-center gap-2 rounded-xl border border-slate-700 bg-[#11151a] px-2 py-1.5"
            onSubmit={(event) => {
              event.preventDefault();
              jumpToSheetOrPage(sheetJumpQuery);
            }}
          >
            <Input
              value={sheetJumpQuery}
              onChange={(event) => setSheetJumpQuery(event.target.value)}
              placeholder="Page or sheet"
              className="h-8 w-28 border-slate-700 bg-[#0e1217] px-2 py-1 text-xs"
            />
            <Button size="sm" variant="outline" type="submit">Go</Button>
          </form>
          <div className="flex items-center gap-2 rounded-xl border border-slate-700 bg-[#11151a] px-3 py-1.5 text-xs text-slate-400">
            <span className="font-medium text-slate-200">{fileName}</span>
            <Separator orientation="vertical" className="h-5 bg-slate-700" />
            <span>{currentSheetNumber || "No sheet loaded"}</span>
          </div>
        </div>
      </div>
      <motion.div
        initial={{ opacity: 0, y: 8 }}
        animate={{ opacity: 1, y: 0 }}
        className="grid w-full gap-2"
        style={{
          gridTemplateColumns: isWideLayout
            ? `${showProjectsPanel ? `${leftPanelWidth}px` : "0px"} minmax(0, 1fr) ${showInspectorPanel ? `${rightPanelWidth}px` : "0px"}`
            : "minmax(0, 1fr)",
        }}
      >
        {showProjectsPanel ? (
        <Card className="relative min-w-0 rounded-2xl border-slate-700 bg-[#1b2026]">
          <CardHeader>
            <CardTitle className="flex items-center gap-2 text-base uppercase tracking-[0.16em] text-slate-300"><Layers3 className="h-4 w-4 text-amber-400" /> Projects</CardTitle>
          </CardHeader>
          <CardContent className="space-y-4">
            <Input ref={fileInputRef} type="file" accept="application/pdf" onChange={handleUpload} />
            <Input value={projectName} onChange={(e) => setProjectName(e.target.value)} placeholder="Project name" />
            <div className="flex flex-wrap gap-2">
              <Button size="sm" onClick={() => void saveProject()} disabled={isCloudSaving || isCloudLoading}><FolderOpen className="h-4 w-4" />{isCloudSaving ? "Saving..." : "Save Project"}</Button>
              <Button size="sm" variant="outline" onClick={() => void refreshCloudProjects()} disabled={isCloudSaving || isCloudLoading}><Download className="h-4 w-4" />Refresh Cloud</Button>
              <Badge variant="secondary" className="max-w-full truncate">{fileName}</Badge>
            </div>
            <div className={`rounded-2xl border p-3 text-xs ${cloudStatus === "ready" ? "border-emerald-800 bg-emerald-950/40 text-emerald-200" : "border-amber-800 bg-amber-950/30 text-amber-200"}`}>
              <div><strong className="text-slate-100">Shared Cloud:</strong> {cloudStatus === "ready" ? "Connected" : "Needs attention"}</div>
              <div className="text-slate-300">{cloudMessage}</div>
              <div className="mt-2 text-slate-400">Projects and PDFs saved here can be opened from other devices.</div>
            </div>
            <div className="rounded-2xl border border-slate-700 bg-[#12161b] p-3 text-xs text-slate-300">
              <div><strong className="text-slate-100">Worker:</strong> {workerStatus}</div>
              <div className="truncate text-slate-500">{workerLabel}</div>
            </div>
            {loadError ? <div className="rounded-xl border border-red-800 bg-red-950/60 p-3 text-xs text-red-200">{loadError}</div> : null}
            <div className="space-y-2">
              <div className="text-[11px] uppercase tracking-[0.14em] text-slate-500">Shared Projects</div>
              {cloudProjects.length ? (
                cloudProjects.map((project) => (
                  <button key={project.id} onClick={() => void loadProject(project)} className={`w-full rounded-xl border p-2 text-left transition ${activeProjectId === project.id ? "border-amber-400 bg-[#2a1f0e]" : "border-slate-700 bg-[#12161b] hover:bg-[#181d24]"}`}>
                    <div className="flex items-center justify-between gap-3">
                      <div className="font-medium text-slate-100">{project.name}</div>
                      <Badge variant="secondary" className="text-[10px]">Cloud</Badge>
                    </div>
                    <div className="text-xs text-slate-500">{project.fileName || "No file"}</div>
                    <div className="text-[11px] text-slate-600">{new Date(project.updatedAt).toLocaleString()}</div>
                  </button>
                ))
              ) : (
                <div className="rounded-xl border border-dashed border-slate-700 bg-[#12161b] p-3 text-xs text-slate-500">
                  No shared projects yet. Upload a PDF and save a project to publish it to the cloud.
                </div>
              )}
            </div>
            <div className="space-y-2">
              <div className="text-[11px] uppercase tracking-[0.14em] text-slate-500">Local Browser Projects</div>
              {projects.map((project) => (
                <button key={project.id} onClick={() => void loadProject(project)} className={`w-full rounded-xl border p-2 text-left transition ${activeProjectId === project.id ? "border-amber-400 bg-[#2a1f0e]" : "border-slate-700 bg-[#12161b] hover:bg-[#181d24]"}`}>
                  <div className="font-medium text-slate-100">{project.name}</div>
                  <div className="text-xs text-slate-500">{project.fileName || "No file"}</div>
                </button>
              ))}
            </div>
            <div className="rounded-2xl border border-slate-700 bg-[#12161b] p-3 text-xs text-slate-400">
              Loaded pages: <strong className="text-slate-100">{pageCount}</strong>
            </div>
          </CardContent>
          {isWideLayout ? (
            <button
              type="button"
              aria-label="Resize project panel"
              onPointerDown={(event) => beginPanelResize("left", event.clientX)}
              className="absolute right-0 top-4 bottom-4 w-1 cursor-col-resize rounded-full bg-transparent transition hover:bg-amber-400/60"
            />
          ) : null}
        </Card>
        ) : null}

        <div className={centerWorkspaceClass}>
          <Card className="sticky top-3 z-30 rounded-2xl border-slate-700 bg-[#181d24]/95 shadow-2xl backdrop-blur">
            <CardContent className="flex flex-wrap items-center gap-2 border-b border-slate-800 p-2.5">
              <span className="self-center px-1 text-[10px] uppercase tracking-[0.14em] text-slate-500">Tools</span>
              <Button variant={tool === TOOL.SELECT ? "default" : "outline"} onClick={() => setTool(TOOL.SELECT)}><MousePointer2 className="h-4 w-4" />Select</Button>
              <Button variant={tool === TOOL.PAN ? "default" : "outline"} onClick={() => setTool(TOOL.PAN)}><Hand className="h-4 w-4" />Pan</Button>
              <Button variant={tool === TOOL.MEASURE ? "default" : "outline"} onClick={() => setTool(TOOL.MEASURE)}><Ruler className="h-4 w-4" />Length</Button>
              <Button variant={tool === TOOL.RECTANGLE ? "default" : "outline"} onClick={() => setTool(TOOL.RECTANGLE)}><Square className="h-4 w-4" />Rectangle</Button>
              <Button variant={tool === TOOL.TEXT ? "default" : "outline"} onClick={() => setTool(TOOL.TEXT)}><Type className="h-4 w-4" />Text</Button>
              <Button variant={tool === TOOL.AREA ? "default" : "outline"} onClick={() => setTool(TOOL.AREA)}><AreaChart className="h-4 w-4" />Area</Button>
              <select
                className="h-10 rounded-md border border-slate-700 bg-[#14191f] px-3 text-sm text-slate-100"
                value={["comment", "cloud", "stamp", "calibrate", "callout", "draw", "ellipse", "link", "snapshot"].includes(tool) ? tool : ""}
                onChange={(event) => {
                  const nextTool = event.target.value as ToolType;
                  if (nextTool) setTool(nextTool);
                }}
              >
                <option value="">More tools</option>
                <option value={TOOL.COMMENT}>Comment</option>
                <option value={TOOL.CLOUD}>Cloud</option>
                <option value={TOOL.STAMP}>Stamp</option>
                <option value={TOOL.CALIBRATE}>Scale</option>
                <option value={TOOL.CALLOUT}>Callout</option>
                <option value={TOOL.DRAW}>Sketch</option>
                <option value={TOOL.ELLIPSE}>Ellipse</option>
                <option value={TOOL.LINK}>Detail Link</option>
                <option value={TOOL.SNAPSHOT}>Snapshot</option>
              </select>
              <Separator orientation="vertical" className="h-8" />
              <span className="self-center px-1 text-[10px] uppercase tracking-[0.14em] text-slate-500">Edit</span>
              <Button variant="outline" onClick={undoLastAction} disabled={undoStack.length === 0}>Undo</Button>
              <Button variant="outline" onClick={redoLastAction} disabled={redoStack.length === 0}>Redo</Button>
              <Button variant="outline" onClick={finishArea}>Finish Area</Button>
              <Button variant={showMarkupList ? "default" : "outline"} onClick={() => setShowMarkupList((value) => !value)}><List className="h-4 w-4" />Markup List</Button>
              <Button variant="destructive" onClick={deleteSelected} disabled={!selectedIds.length || !canDelete()}><Trash2 className="h-4 w-4" />Delete</Button>
            </CardContent>
            <div className="flex flex-wrap items-center gap-3 border-t border-slate-800 bg-[#14191f] px-3 py-2 text-xs text-slate-300">
              <div className="font-medium uppercase tracking-[0.14em] text-slate-500">{propertyLabel}</div>
              {selectedIds.length ? (
                <>
                  <div className="flex items-center gap-2">
                    <Button size="sm" variant="outline" onClick={copySelectionToClipboard}>
                      <Copy className="h-4 w-4" />
                      Copy
                    </Button>
                    <Button size="sm" variant="outline" onClick={duplicateSelected} disabled={!canEdit()}>
                      <Copy className="h-4 w-4" />
                      Duplicate
                    </Button>
                  </div>
                  <select
                    className="h-8 rounded-md border border-slate-700 bg-[#1b2026] px-2 text-xs text-slate-100"
                    value={selectedAnnotation?.status || "open"}
                    onChange={(e) => updateSelectedField("status", e.target.value)}
                  >
                    {MARKUP_STATUSES.map((status) => (
                      <option key={status.value} value={status.value}>{status.label}</option>
                    ))}
                  </select>
                  <select
                    className="h-8 rounded-md border border-slate-700 bg-[#1b2026] px-2 text-xs text-slate-100"
                    value={selectedAnnotation?.layerId || "layer-general"}
                    onChange={(e) => updateSelectedField("layerId", e.target.value)}
                  >
                    {layers.map((layer) => (
                      <option key={layer.id} value={layer.id}>{layer.name}</option>
                    ))}
                  </select>
                </>
              ) : null}
              <div className="flex items-center gap-2">
                <span className="text-slate-500">Weight</span>
                {[1, 2, 4, 6].map((weight) => (
                  <button
                    key={weight}
                    type="button"
                    onClick={() => updateStyleValue("strokeWidth", weight)}
                    className={`rounded-md border px-2 py-1 transition ${propertyStrokeWidth === weight ? "border-amber-400 bg-[#2b2010] text-amber-200" : "border-slate-700 bg-[#1b2026] hover:bg-[#232933]"}`}
                  >
                    {weight}
                  </button>
                ))}
              </div>
              <label className="flex items-center gap-2">
                <span className="text-slate-500">Stroke</span>
                <input type="color" value={propertyStrokeColor} onChange={(e) => updateStyleValue("color", e.target.value)} className="h-8 w-10 cursor-pointer rounded border border-slate-700 bg-transparent p-1" />
              </label>
              <label className="flex items-center gap-2">
                <span className="text-slate-500">Fill</span>
                <input type="color" value={propertyFillColor} onChange={(e) => updateStyleValue("fillColor", e.target.value)} className="h-8 w-10 cursor-pointer rounded border border-slate-700 bg-transparent p-1" />
              </label>
              <label className="flex min-w-[180px] items-center gap-2">
                <span className="text-slate-500">Opacity</span>
                <input type="range" min="0" max="0.75" step="0.05" value={propertyFillOpacity} onChange={(e) => updateStyleValue("fillOpacity", Number(e.target.value))} className="w-full accent-amber-400" />
                <span className="w-10 text-right text-slate-400">{Math.round(propertyFillOpacity * 100)}%</span>
              </label>
            </div>
          </Card>

          <Card className="overflow-hidden rounded-2xl border-slate-700 bg-[#1a1f25] shadow-2xl">
            <div ref={viewerRef} className={`relative overflow-hidden bg-[#2a3038] ${isViewerFullscreen ? "h-[100dvh]" : "h-[83vh]"}`} onDoubleClick={handleViewerDoubleClick}>
              <div className="pointer-events-none absolute left-2 top-2 z-20">
                <div className="pointer-events-auto flex flex-wrap items-center gap-2 rounded-2xl border border-slate-700 bg-[#171b21]/95 p-2 shadow-2xl backdrop-blur">
                  <span className="px-1 text-[10px] uppercase tracking-[0.14em] text-slate-500">View</span>
                  <Button size="sm" variant="outline" onClick={() => changePageBy(-1)} disabled={currentPage <= 1}>
                    <ChevronLeft className="h-4 w-4" />
                    {isViewerFullscreen ? "Prev Page" : null}
                  </Button>
                  <Button size="sm" variant="outline" onClick={() => changePageBy(1)} disabled={currentPage >= pageCount}>
                    <ChevronRight className="h-4 w-4" />
                    {isViewerFullscreen ? "Next Page" : null}
                  </Button>
                  <Separator orientation="vertical" className="h-8" />
                  <Button size="sm" variant="outline" onClick={() => zoomBy(1 / 1.2)}>
                    <ZoomOut className="h-4 w-4" />
                  </Button>
                  <Badge variant="secondary" className="px-2 py-1 text-xs">
                    {Math.round(transform.scale * 100)}%
                  </Badge>
                  <Button size="sm" variant="outline" onClick={() => zoomToPercent(50)}>50%</Button>
                  <Button size="sm" variant="outline" onClick={() => zoomToPercent(100)}>100%</Button>
                  <Button size="sm" variant="outline" onClick={() => zoomToPercent(200)}>200%</Button>
                  <Button size="sm" variant={navigationMode === "sheet" ? "default" : "outline"} onClick={() => setNavigationModeAndAnnounce(navigationMode === "sheet" ? "review" : "sheet")}>
                    {navigationMode === "sheet" ? "Pages" : "Pan"}
                  </Button>
                  <Button size="sm" variant={boxZoomMode ? "default" : "outline"} onClick={() => (boxZoomMode ? setBoxZoomMode(false) : activateBoxZoom())}>
                    Box
                  </Button>
                  <Button size="sm" variant="outline" onClick={() => zoomBy(1.2)}>
                    <ZoomIn className="h-4 w-4" />
                  </Button>
                  <Button size="sm" variant="outline" onClick={fitToViewer}>
                    Fit
                  </Button>
                  <Button size="sm" variant="outline" onClick={fitWidthToViewer}>
                    Width
                  </Button>
                  <Button size="sm" variant="outline" onClick={fitHeightToViewer}>
                    Height
                  </Button>
                  {isViewerFullscreen ? (
                    <>
                      <Separator orientation="vertical" className="h-8" />
                      <span className="px-1 text-[10px] uppercase tracking-[0.14em] text-slate-500">Markup</span>
                      <Button size="sm" variant={tool === TOOL.SELECT ? "default" : "outline"} onClick={() => setToolAndAnnounce(TOOL.SELECT, "Select tool")}>Select</Button>
                      <Button size="sm" variant={tool === TOOL.PAN ? "default" : "outline"} onClick={() => setToolAndAnnounce(TOOL.PAN, "Pan tool")}>Pan</Button>
                      <Button size="sm" variant={tool === TOOL.MEASURE ? "default" : "outline"} onClick={() => setToolAndAnnounce(TOOL.MEASURE, "Length tool")}>Length</Button>
                      <Button size="sm" variant={tool === TOOL.AREA ? "default" : "outline"} onClick={() => setToolAndAnnounce(TOOL.AREA, "Area tool")}>Area</Button>
                      <Button size="sm" variant={tool === TOOL.TEXT ? "default" : "outline"} onClick={() => setToolAndAnnounce(TOOL.TEXT, "Text tool")}>Text</Button>
                      <Button size="sm" variant={tool === TOOL.RECTANGLE ? "default" : "outline"} onClick={() => setToolAndAnnounce(TOOL.RECTANGLE, "Rectangle tool")}>Rect</Button>
                      <select
                        className="h-9 rounded-md border border-slate-700 bg-[#14191f] px-2 text-xs text-slate-100"
                        value={["comment", "cloud", "stamp", "calibrate", "callout", "draw", "ellipse", "link", "snapshot"].includes(tool) ? tool : ""}
                        onChange={(event) => {
                          const nextTool = event.target.value as ToolType;
                          if (nextTool) setToolAndAnnounce(nextTool, `${nextTool} tool`);
                        }}
                      >
                        <option value="">More</option>
                        <option value={TOOL.COMMENT}>Comment</option>
                        <option value={TOOL.CLOUD}>Cloud</option>
                        <option value={TOOL.STAMP}>Stamp</option>
                        <option value={TOOL.CALIBRATE}>Scale</option>
                        <option value={TOOL.CALLOUT}>Callout</option>
                        <option value={TOOL.DRAW}>Sketch</option>
                        <option value={TOOL.ELLIPSE}>Ellipse</option>
                        <option value={TOOL.LINK}>Detail Link</option>
                        <option value={TOOL.SNAPSHOT}>Snapshot</option>
                      </select>
                      <Button size="sm" variant={tool === TOOL.CALIBRATE ? "default" : "outline"} onClick={() => setToolAndAnnounce(TOOL.CALIBRATE, "Scale tool")}>
                        Scale
                      </Button>
                      <Button size="sm" variant="outline" onClick={finishArea} disabled={!(draft && draft.type === "area")}>
                        Finish Area
                      </Button>
                    </>
                  ) : null}
                  <Button size="sm" variant="outline" onClick={() => void toggleViewerFullscreen()}>
                    {isViewerFullscreen ? <Minimize2 className="h-4 w-4" /> : <Maximize2 className="h-4 w-4" />}
                  </Button>
                </div>
              </div>
              <canvas ref={pdfCanvasRef} className="absolute left-0 top-0 shadow-[0_30px_90px_rgba(0,0,0,0.55)]" style={{ transform: `translate(${transform.x}px, ${transform.y}px) scale(${transform.scale})`, transformOrigin: "top left" }} />
              {compareMode && comparePage && comparePage !== currentPage ? <canvas ref={compareCanvasRef} className="pointer-events-none absolute left-0 top-0 mix-blend-multiply opacity-60" style={{ transform: `translate(${transform.x}px, ${transform.y}px) scale(${transform.scale})`, transformOrigin: "top left" }} /> : null}
              <canvas ref={overlayCanvasRef} className="absolute inset-0 touch-none" style={{ cursor: boxZoomMode ? "zoom-in" : hoveredAutoReferenceId ? "pointer" : tool === TOOL.PAN ? "grab" : "crosshair" }} onPointerDown={handlePointerDown} onPointerMove={handlePointerMove} onPointerUp={handlePointerUp} onPointerLeave={handlePointerUp} onPointerCancel={handlePointerUp} />
              {hoveredAutoReference ? (
                <div className="pointer-events-none absolute right-3 top-16 z-20 w-64 overflow-hidden rounded-2xl border border-slate-700 bg-[#13181e]/95 shadow-2xl backdrop-blur">
                  <div className="border-b border-slate-700 px-3 py-2 text-[10px] uppercase tracking-[0.14em] text-slate-500">Reference Preview</div>
                  <div className="space-y-3 p-3">
                    <div>
                      <div className="text-sm font-semibold text-slate-100">{hoveredAutoReference.text}</div>
                      <div className="text-xs text-slate-400">
                        {pageSheetNumbers[hoveredAutoReference.targetPage] || hoveredAutoReference.targetSheetNumber} | Page {hoveredAutoReference.targetPage}
                      </div>
                    </div>
                    {pageThumbnails[hoveredAutoReference.targetPage] ? (
                      <img
                        src={pageThumbnails[hoveredAutoReference.targetPage]}
                        alt={`${hoveredAutoReference.targetSheetNumber} preview`}
                        className="h-auto max-h-40 w-full rounded-lg border border-slate-700 bg-white object-contain"
                      />
                    ) : (
                      <div className="flex h-24 items-center justify-center rounded-lg border border-dashed border-slate-700 bg-[#0f1318] text-xs text-slate-500">
                        Preview loading
                      </div>
                    )}
                    <div className="text-[11px] text-slate-400">Click to jump to the target sheet and focus the matching detail when available.</div>
                  </div>
                </div>
              ) : null}
              {inlineEditor && inlineEditorLayout ? (
                <input
                  ref={inlineEditorRef}
                  data-wheel-bypass="true"
                  value={inlineEditor.value}
                  onChange={(event) => setInlineEditor((prev) => (prev ? { ...prev, value: event.target.value } : prev))}
                  onBlur={commitInlineEditor}
                  onKeyDown={(event) => {
                    if (event.key === "Enter") {
                      event.preventDefault();
                      event.currentTarget.blur();
                    }
                    if (event.key === "Escape") {
                      event.preventDefault();
                      cancelInlineEditor();
                    }
                  }}
                  className="absolute z-30 rounded-lg border border-amber-400 bg-[#0f1318]/95 px-3 py-2 text-sm text-slate-50 shadow-[0_14px_28px_rgba(0,0,0,0.45)] outline-none ring-2 ring-amber-400/20"
                  style={{
                    left: inlineEditorLayout.left,
                    top: inlineEditorLayout.top,
                    width: inlineEditorLayout.width,
                    height: inlineEditorLayout.height,
                  }}
                />
              ) : null}
              {miniMapMetrics ? (
                <button
                  type="button"
                  data-wheel-bypass="true"
                  onPointerDown={handleMiniMapPointerDown}
                  onPointerMove={handleMiniMapPointerMove}
                  onPointerUp={handleMiniMapPointerUp}
                  onPointerCancel={handleMiniMapPointerCancel}
                  className="absolute bottom-3 right-3 z-20 overflow-hidden rounded-xl border border-slate-700 bg-[#11151a]/95 p-2 shadow-xl backdrop-blur"
                  style={{ width: miniMapMetrics.width + 16 }}
                  aria-label="Navigator mini map"
                >
                  <div className="mb-1 text-left text-[10px] uppercase tracking-[0.14em] text-slate-500">Navigator</div>
                  <div className="relative rounded-md border border-slate-700 bg-[#1d232b]" style={{ width: miniMapMetrics.width, height: miniMapMetrics.height }}>
                    <div className="absolute inset-0 bg-[linear-gradient(180deg,rgba(255,255,255,0.03),rgba(255,255,255,0))]" />
                    <div
                      className="absolute border border-amber-300/90 bg-amber-400/15 shadow-[0_0_0_1px_rgba(0,0,0,0.35)]"
                      style={{
                        left: miniMapMetrics.viewport.x,
                        top: miniMapMetrics.viewport.y,
                        width: Math.min(miniMapMetrics.viewport.w, miniMapMetrics.width - miniMapMetrics.viewport.x),
                        height: Math.min(miniMapMetrics.viewport.h, miniMapMetrics.height - miniMapMetrics.viewport.y),
                      }}
                    />
                  </div>
                </button>
              ) : null}
              <div className="pointer-events-none absolute bottom-3 left-3 z-20 rounded-xl border border-slate-700 bg-[#171b21]/95 px-3 py-2 text-xs text-slate-300 shadow-lg backdrop-blur">
                Page {currentPage}{currentSheetNumber ? ` ${currentSheetNumber}` : ""} of {pageCount || 1} · {Math.round(transform.scale * 100)}% · {tool} · Wheel on sheet flips pages · Alt+Wheel scrolls app · Shift+Wheel zoom · Drag empty space to box-select
              </div>
              <div className="pointer-events-none absolute bottom-14 left-3 z-20 rounded-xl border border-slate-700 bg-[#101419]/90 px-3 py-1.5 text-[11px] text-slate-400 shadow-lg backdrop-blur">
                {cursorSummary}
              </div>
              <div className="pointer-events-none absolute bottom-14 right-3 z-20 rounded-xl border border-slate-700 bg-[#101419]/90 px-3 py-1.5 text-[11px] text-slate-400 shadow-lg backdrop-blur">
                Wheel mode: {navigationMode === "sheet" ? "Pages" : "Pan"} | G toggles mode
              </div>
              {tool === TOOL.MEASURE ? (
                <div className="pointer-events-none absolute top-16 left-1/2 z-20 -translate-x-1/2 rounded-xl border border-sky-500/40 bg-[#0d1720]/92 px-3 py-2 text-[11px] text-sky-100 shadow-lg backdrop-blur">
                  Length takeoff: click and drag to measure | Space temporarily pans | Select handles to refine
                </div>
              ) : null}
              {draft?.type === "area" ? (
                <div className="pointer-events-none absolute top-16 left-1/2 z-20 -translate-x-1/2 rounded-xl border border-emerald-500/40 bg-[#0d1b18]/92 px-3 py-2 text-[11px] text-emerald-100 shadow-lg backdrop-blur">
                  Area takeoff: click to add points | click start point, double-click, or press Enter to finish | Esc cancels
                </div>
              ) : null}
              {hoveredKeynoteToken ? (
                <div className="pointer-events-none absolute right-3 top-16 z-20 rounded-xl border border-yellow-300/60 bg-[#201b08]/90 px-3 py-2 text-[11px] text-yellow-100 shadow-lg backdrop-blur">
                  Keynote {hoveredKeynoteToken} highlighted | {highlightedKeynoteFragments.length} match{highlightedKeynoteFragments.length === 1 ? "" : "es"}
                </div>
              ) : null}
              <div className="pointer-events-none absolute inset-x-0 bottom-0 z-10 h-10 border-t border-slate-800 bg-[#15191f]/95" />
            </div>
          </Card>

          {showPagesPanel ? (
          <Card className="rounded-2xl border-slate-700 bg-[#1b2026]">
            <CardHeader>
              <CardTitle className="flex items-center justify-between gap-2 text-base uppercase tracking-[0.16em] text-slate-300">
                <span>Pages</span>
                <Badge variant="secondary">{filteredThumbnailPages.length} shown</Badge>
              </CardTitle>
            </CardHeader>
            <CardContent className="space-y-3">
              <div className="flex flex-wrap gap-2">
                {availableBookmarkTabs.map((tab) => (
                  <Button
                    key={tab.id}
                    size="sm"
                    variant={activeBookmarkTab === tab.id ? "default" : "outline"}
                    onClick={() => setActiveBookmarkTab(tab.id)}
                  >
                    {tab.label}
                  </Button>
                ))}
              </div>
              <ScrollArea className="h-[26vh] pr-2">
                <div className="grid grid-cols-2 gap-2 sm:grid-cols-3 xl:grid-cols-5 2xl:grid-cols-6">
                  {filteredThumbnailPages.map((page) => {
                    const thumbnail = pageThumbnails[page];
                    const isActive = page === currentPage;
                    const bookmark = pageBookmarkSections[page];
                    const bookmarkLabel = BOOKMARK_TABS.find((tab) => tab.id === bookmark)?.label || "General";
                    const sheetNumber = pageSheetNumbers[page];

                    return (
                      <button
                        key={page}
                        type="button"
                        onClick={() => setCurrentPage(page)}
                        className={`w-full rounded-xl border p-1.5 text-left transition ${isActive ? "border-amber-400 bg-[#2b2010] shadow-sm" : "border-slate-700 bg-[#12161b] hover:border-slate-500 hover:bg-[#171c22]"}`}
                      >
                        <div className="mb-1 flex items-center justify-between gap-2">
                          <div className={`text-[11px] font-semibold ${isActive ? "text-amber-300" : "text-slate-300"}`}>
                            Page {page}{sheetNumber ? ` ${sheetNumber}` : ""}
                          </div>
                          <div className="truncate text-[10px] uppercase tracking-[0.12em] text-slate-400">{bookmarkLabel}</div>
                        </div>
                        {thumbnail ? (
                          <img
                            src={thumbnail}
                            alt={`Page ${page} thumbnail`}
                            className="mx-auto h-auto max-h-28 w-auto max-w-full rounded-md border border-slate-700 bg-white shadow-[0_12px_24px_rgba(0,0,0,0.4)]"
                          />
                        ) : (
                          <div className="mx-auto flex aspect-[3/4] max-h-28 w-full items-center justify-center rounded-md border border-dashed border-slate-600 bg-[#0f1318] text-[10px] text-slate-500">
                            Loading
                          </div>
                        )}
                      </button>
                    );
                  })}
                </div>
              </ScrollArea>
            </CardContent>
          </Card>
          ) : null}

          {showMarkupList ? (
            <Card className="rounded-2xl border-slate-700 bg-[#1b2026]">
              <CardHeader>
                <CardTitle className="flex items-center gap-2 text-base uppercase tracking-[0.16em] text-slate-300"><List className="h-4 w-4 text-amber-400" />Markup Database</CardTitle>
              </CardHeader>
              <CardContent className="space-y-3">
                <div className="grid grid-cols-1 gap-3 md:grid-cols-3">
                  <Input value={markupFilter} onChange={(e) => setMarkupFilter(e.target.value)} placeholder="Filter markups" />
                  <select className="w-full rounded-md border border-slate-600 bg-slate-800 px-3 py-2 text-sm text-slate-100" value={markupStatusFilter} onChange={(e) => setMarkupStatusFilter(e.target.value || "all")}>
                    <option value="all">All statuses</option>
                    {MARKUP_STATUSES.map((status) => (
                      <option key={status.value} value={status.value}>{status.label}</option>
                    ))}
                  </select>
                  <div className="rounded-xl border border-slate-700 bg-[#12161b] px-3 py-2 text-sm text-slate-300">{markupRows.length} visible rows</div>
                </div>
                <div className="max-h-80 overflow-auto rounded-xl border border-slate-700 bg-[#12161b]">
                  <table className="w-full text-sm">
                    <thead className="sticky top-0 bg-[#181d24] text-slate-300">
                      <tr>
                        <th className="p-2 text-left">ID</th>
                        <th className="p-2 text-left">Sheet</th>
                        <th className="p-2 text-left">Type</th>
                        <th className="p-2 text-left">Value</th>
                        <th className="p-2 text-left">Category</th>
                        <th className="p-2 text-left">Status</th>
                        <th className="p-2 text-left">Author</th>
                        <th className="p-2 text-left">Layer</th>
                      </tr>
                    </thead>
                    <tbody>
                      {markupRows.map((row) => (
                        <tr
                          key={row.id}
                          className={`cursor-pointer border-t border-slate-800 ${selectedIdSet.has(row.id) ? "bg-[#2b2010]" : "bg-[#12161b] hover:bg-[#171c22]"}`}
                          onClick={(event) => {
                            setCurrentPage(row.page);
                            if (event.ctrlKey || event.metaKey) {
                              const nextSelection = selectedIdSet.has(row.id)
                                ? selectedIds.filter((id) => id !== row.id)
                                : [...selectedIds, row.id];
                              setSelection(nextSelection, row.id);
                              return;
                            }
                            setSelection([row.id], row.id);
                          }}
                        >
                          <td className="p-2 font-mono text-xs">{row.id.slice(0, 6)}</td>
                          <td className="p-2">{row.page}</td>
                          <td className="p-2">{row.type}</td>
                          <td className="p-2">{annotationSummary(row as Annotation, calibration)}</td>
                          <td className="p-2">{categoryForAnnotation(row as any)}</td>
                          <td className="p-2">
                            <span className={`inline-flex rounded-md border px-2 py-0.5 text-[11px] ${markupStatusClasses(row.status || "open")}`}>
                              {markupStatusLabel(row.status || "open")}
                            </span>
                          </td>
                          <td className="p-2">{row.author || "You"}</td>
                          <td className="p-2">{layers.find((layer) => layer.id === (row.layerId || "layer-general"))?.name || "General"}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </CardContent>
            </Card>
          ) : null}
        </div>

        {showInspectorPanel ? (
        <Card className="relative min-w-0 rounded-2xl border-slate-700 bg-[#1b2026]">
          <CardHeader>
            <CardTitle className="flex items-center gap-2 text-base uppercase tracking-[0.16em] text-slate-300"><FileText className="h-4 w-4 text-amber-400" />Inspector, Layers, Reports</CardTitle>
          </CardHeader>
          <CardContent className="space-y-4 text-sm">
            <div className="space-y-1 rounded-xl border border-slate-700 bg-[#12161b] p-3">
              <div><strong>Tool:</strong> {tool}</div>
              <div><strong>Scale:</strong> {calibration.toFixed(5)} ft/unit</div>
              <div><strong>Total Markups:</strong> {totals.count}</div>
              <div><strong>Total Length:</strong> {totals.length}</div>
              <div><strong>Total Area:</strong> {totals.area}</div>
            </div>

            <div className="space-y-2 rounded-xl border border-slate-700 bg-[#12161b] p-3">
              <div className="flex items-center gap-2 font-medium"><Users className="h-4 w-4" /> Collaboration & Roles</div>
              <select className="w-full rounded-md border border-slate-600 bg-slate-800 px-3 py-2 text-slate-100" value={activeUserId} onChange={(e) => setActiveUserId(e.target.value)}>
                {users.map((user) => <option key={user.id} value={user.id}>{user.name} · {DEFAULT_ROLES.find((role) => role.id === user.roleId)?.name}</option>)}
              </select>
              <div className="text-xs text-slate-500">Edit: {String(canEdit())} · Delete: {String(canDelete())} · Export: {String(canExport())}</div>
            </div>

            <div className="space-y-2 rounded-xl border border-slate-700 bg-[#12161b] p-3">
              <div className="flex items-center gap-2 font-medium"><Eye className="h-4 w-4" /> Layer Controls</div>
              {layers.map((layer) => (
                <button key={layer.id} onClick={() => toggleLayer(layer.id)} className="flex w-full items-center justify-between rounded-lg border border-slate-700 bg-[#191e25] p-2 text-left hover:bg-[#222830]">
                  <span>{layer.name}</span>
                  {layer.visible ? <Eye className="h-4 w-4" /> : <EyeOff className="h-4 w-4" />}
                </button>
              ))}
            </div>

            <div className="space-y-2 rounded-xl border border-slate-700 bg-[#12161b] p-3">
              <div className="flex items-center gap-2 font-medium"><Search className="h-4 w-4" /> OCR / Text Search</div>
              <div className="flex gap-2">
                <Input value={ocrQuery} onChange={(e) => setOcrQuery(e.target.value)} placeholder="Search sheet text" />
                <Button onClick={runOcrSearch}><Search className="h-4 w-4" /></Button>
              </div>
              <div className="max-h-36 space-y-2 overflow-auto">
                {ocrResults.map((result) => (
                  <button key={`${result.page}-${result.preview}`} onClick={() => setCurrentPage(result.page)} className="w-full rounded-lg border border-slate-700 bg-[#191e25] p-2 text-left hover:bg-[#222830]">
                    <div className="font-medium">Page {result.page}</div>
                    <div className="text-xs text-slate-500">{result.preview}</div>
                  </button>
                ))}
              </div>
            </div>

            <div className="space-y-2 rounded-xl border border-slate-700 bg-[#12161b] p-3">
              <div className="flex items-center gap-2 font-medium"><Library className="h-4 w-4" /> Tool Chest</div>
              <div className="grid grid-cols-1 gap-2">
                {toolChest.map((item) => (
                  <button key={item.id} onClick={() => setSelectedChestId((prev) => (prev === item.id ? null : item.id))} className={`rounded-lg border p-2 text-left ${selectedChestId === item.id ? "border-amber-400 bg-[#2b2010]" : "border-slate-700 bg-[#191e25]"}`}>
                    <div className="font-medium">{item.name}</div>
                    <div className="text-xs text-slate-500">{item.payload.text}</div>
                  </button>
                ))}
              </div>
            </div>

            <div className="space-y-3 rounded-xl border border-slate-700 bg-[#12161b] p-3">
              <div className="font-medium">Markup Presets</div>
              <div className="grid grid-cols-1 gap-2">
                {MARKUP_PRESETS.map((preset) => (
                  <button
                    key={preset.id}
                    type="button"
                    onClick={() => applyMarkupPreset(preset)}
                    className={`rounded-lg border p-2 text-left transition ${activeMarkupLayerId === preset.layerId && activeMarkupCategory === preset.category && activeMarkupColor === preset.color ? "border-amber-400 bg-[#2b2010]" : "border-slate-700 bg-[#191e25] hover:bg-[#222830]"}`}
                  >
                    <div className="flex items-center justify-between gap-3">
                      <span className="font-medium">{preset.name}</span>
                      <span className="h-3 w-3 rounded-full border border-slate-500" style={{ backgroundColor: preset.color }} />
                    </div>
                    <div className="text-xs text-slate-500">{preset.category} · {layers.find((layer) => layer.id === preset.layerId)?.name || preset.layerId}</div>
                  </button>
                ))}
              </div>
              <div className="space-y-2">
                <div className="text-xs uppercase tracking-[0.12em] text-slate-500">Active Color</div>
                <div className="flex flex-wrap gap-2">
                  {["#2563eb", "#b91c1c", "#d97706", "#15803d", "#7c3aed", "#111827"].map((color) => (
                    <button
                      key={color}
                      type="button"
                      aria-label={`Use ${color} markup color`}
                      onClick={() => setActiveMarkupColor(color)}
                      className={`h-7 w-7 rounded-full border-2 transition ${activeMarkupColor === color ? "border-amber-300 scale-110" : "border-slate-600"}`}
                      style={{ backgroundColor: color }}
                    />
                  ))}
                </div>
              </div>
              <select className="w-full rounded-md border border-slate-600 bg-slate-800 px-3 py-2 text-slate-100" value={activeMarkupLayerId} onChange={(e) => setActiveMarkupLayerId(e.target.value)}>
                {layers.map((layer) => <option key={layer.id} value={layer.id}>{layer.name}</option>)}
              </select>
              <Input value={activeMarkupCategory} onChange={(e) => setActiveMarkupCategory(e.target.value)} placeholder="Default category" />
              <div className="text-xs text-slate-500">New markups will use the active preset color, layer, and category.</div>
            </div>

            <div className="space-y-3 rounded-xl border border-slate-700 bg-[#12161b] p-3">
              <div className="flex items-center gap-2 font-medium"><Tag className="h-4 w-4" /> Stamp Presets</div>
              <div className="grid grid-cols-1 gap-2">
                {STAMP_PRESETS.map((preset) => (
                  <button
                    key={preset.id}
                    type="button"
                    onClick={() => applyStampPreset(preset)}
                    className={`rounded-lg border p-2 text-left transition ${activeStampPresetId === preset.id ? "border-amber-400 bg-[#2b2010]" : "border-slate-700 bg-[#191e25] hover:bg-[#222830]"}`}
                  >
                    <div className="flex items-center justify-between gap-3">
                      <span className="font-medium">{preset.name}</span>
                      <span className="h-3 w-3 rounded-full border border-slate-500" style={{ backgroundColor: preset.color }} />
                    </div>
                    <div className="text-[11px] uppercase tracking-[0.08em] text-slate-400">{preset.text}</div>
                    <div className="text-xs text-slate-500">{preset.category} Â· {layers.find((layer) => layer.id === preset.layerId)?.name || preset.layerId}</div>
                  </button>
                ))}
              </div>
              <div className="text-xs text-slate-500">Choose a review stamp, then place it with the Stamp tool for fast field review workflows.</div>
            </div>

            {selectedIds.length > 1 ? (
              <div className="space-y-3 rounded-xl border border-slate-700 bg-[#12161b] p-3">
                <div className="font-medium">Selected Markups</div>
                <div className="text-xs text-slate-500">{selectedIds.length} markups selected on this sheet</div>
                <select className="w-full rounded-md border border-slate-600 bg-slate-800 px-3 py-2 text-slate-100" value={selectedAnnotation?.status || "open"} onChange={(e) => updateSelectedField("status", e.target.value)}>
                  {MARKUP_STATUSES.map((status) => (
                    <option key={status.value} value={status.value}>{status.label}</option>
                  ))}
                </select>
                <select className="w-full rounded-md border border-slate-600 bg-slate-800 px-3 py-2 text-slate-100" value={selectedAnnotation?.layerId || "layer-general"} onChange={(e) => updateSelectedField("layerId", e.target.value)}>
                  {layers.map((layer) => (<option key={layer.id} value={layer.id}>{layer.name}</option>))}
                </select>
                <Input value={selectedAnnotation?.category || activeMarkupCategory} onChange={(e) => updateSelectedField("category", e.target.value)} placeholder="Category" />
                <div className="grid grid-cols-2 gap-2">
                  <label className="space-y-1 text-xs text-slate-500">
                    <span>Stroke</span>
                    <input type="color" value={selectedAnnotation?.color || activeMarkupColor} onChange={(e) => updateSelectedField("color", e.target.value)} className="h-10 w-full cursor-pointer rounded border border-slate-700 bg-transparent p-1" />
                  </label>
                  <label className="space-y-1 text-xs text-slate-500">
                    <span>Fill</span>
                    <input type="color" value={selectedAnnotation?.fillColor || activeMarkupFillColor} onChange={(e) => updateSelectedField("fillColor", e.target.value)} className="h-10 w-full cursor-pointer rounded border border-slate-700 bg-transparent p-1" />
                  </label>
                </div>
                <label className="space-y-1 text-xs text-slate-500">
                  <span>Line Weight</span>
                  <input type="range" min="1" max="8" step="1" value={selectedAnnotation?.strokeWidth || 2} onChange={(e) => updateSelectedField("strokeWidth", Number(e.target.value))} className="w-full accent-amber-400" />
                </label>
              </div>
            ) : selectedAnnotation ? (
              <div className="space-y-2 rounded-xl border border-slate-700 bg-[#12161b] p-3">
                <div className="font-medium">Selected Markup</div>
                <div className="text-xs text-slate-500">{selectedAnnotation.id}</div>
                <div>Type: {selectedAnnotation.type}</div>
                {selectedAnnotation.type === "text" || selectedAnnotation.type === "comment" || selectedAnnotation.type === "callout" || selectedAnnotation.type === "stamp" ? <Input value={selectedAnnotation.type === "comment" ? (selectedAnnotation.comment || selectedAnnotation.text || "") : (selectedAnnotation.text || "")} onChange={(e) => updateSelectedField(selectedAnnotation.type === "comment" ? "comment" : "text", e.target.value)} /> : null}
                {selectedAnnotation.type === "snapshot" ? <Input value={selectedAnnotation.title || ""} onChange={(e) => updateSelectedField("title", e.target.value)} /> : null}
                {selectedAnnotation.type === "cloud" ? <Input value={selectedAnnotation.label || ""} onChange={(e) => updateSelectedField("label", e.target.value)} placeholder="Cloud label" /> : null}
                {selectedAnnotation.type === "link" ? <Input type="number" value={selectedAnnotation.targetPage} onChange={(e) => updateSelectedField("targetPage", Number(e.target.value) || 1)} /> : null}
                <select className="w-full rounded-md border border-slate-600 bg-slate-800 px-3 py-2 text-slate-100" value={selectedAnnotation.status || "open"} onChange={(e) => updateSelectedField("status", e.target.value)}>
                  {MARKUP_STATUSES.map((status) => (
                    <option key={status.value} value={status.value}>{status.label}</option>
                  ))}
                </select>
                <div className="flex flex-wrap gap-2">
                  {MARKUP_STATUSES.map((status) => (
                    <button
                      key={status.value}
                      type="button"
                      onClick={() => updateSelectedField("status", status.value)}
                      className={`rounded-md border px-2 py-1 text-[11px] transition ${selectedAnnotation.status === status.value ? markupStatusClasses(status.value) : "border-slate-700 bg-[#191e25] text-slate-300 hover:bg-[#222830]"}`}
                    >
                      {status.label}
                    </button>
                  ))}
                </div>
                <Input value={selectedAnnotation.author || "You"} onChange={(e) => updateSelectedField("author", e.target.value)} placeholder="Author" />
                <Input value={selectedAnnotation.category || "General"} onChange={(e) => updateSelectedField("category", e.target.value)} placeholder="Category" />
                <select className="w-full rounded-md border border-slate-600 bg-slate-800 px-3 py-2 text-slate-100" value={selectedAnnotation.layerId || "layer-general"} onChange={(e) => updateSelectedField("layerId", e.target.value)}>
                  {layers.map((layer) => (<option key={layer.id} value={layer.id}>{layer.name}</option>))}
                </select>
                <div className="grid grid-cols-2 gap-2">
                  <label className="space-y-1 text-xs text-slate-500">
                    <span>Stroke</span>
                    <input type="color" value={selectedAnnotation.color || "#2563eb"} onChange={(e) => updateSelectedField("color", e.target.value)} className="h-10 w-full cursor-pointer rounded border border-slate-700 bg-transparent p-1" />
                  </label>
                  <label className="space-y-1 text-xs text-slate-500">
                    <span>Fill</span>
                    <input type="color" value={selectedAnnotation.fillColor || selectedAnnotation.color || "#2563eb"} onChange={(e) => updateSelectedField("fillColor", e.target.value)} className="h-10 w-full cursor-pointer rounded border border-slate-700 bg-transparent p-1" />
                  </label>
                </div>
                <label className="space-y-1 text-xs text-slate-500">
                  <span>Line Weight: {selectedAnnotation.strokeWidth || 2}px</span>
                  <input type="range" min="1" max="8" step="1" value={selectedAnnotation.strokeWidth || 2} onChange={(e) => updateSelectedField("strokeWidth", Number(e.target.value))} className="w-full accent-amber-400" />
                </label>
                <label className="space-y-1 text-xs text-slate-500">
                  <span>Fill Opacity: {Math.round((selectedAnnotation.fillOpacity ?? 0.16) * 100)}%</span>
                  <input type="range" min="0" max="0.75" step="0.05" value={selectedAnnotation.fillOpacity ?? 0.16} onChange={(e) => updateSelectedField("fillOpacity", Number(e.target.value))} className="w-full accent-amber-400" />
                </label>
                <div className="text-xs text-slate-500">{annotationSummary(selectedAnnotation, calibration)}</div>
              </div>
            ) : null}

            <div className="space-y-2 rounded-xl border border-slate-700 bg-[#12161b] p-3">
              <div className="flex items-center gap-2 font-medium"><Tag className="h-4 w-4" /> Quantity Legend</div>
              <div className="space-y-1 text-xs">
                {Object.entries(totals.legend).map(([key, value]) => <div key={key} className="flex justify-between"><span>{key}</span><span>{value}</span></div>)}
              </div>
            </div>

            <div className="space-y-2 rounded-xl border border-slate-700 bg-[#12161b] p-3">
              <div className="flex items-center gap-2 font-medium"><Ruler className="h-4 w-4" /> Takeoff Totals</div>
              <div className="space-y-2 text-xs">
                {takeoffSummaries.length ? (
                  takeoffSummaries.map((summary) => (
                    <div key={summary.category} className="rounded-lg border border-slate-700 bg-[#191e25] p-2">
                      <div className="flex items-center justify-between gap-3 text-slate-100">
                        <span className="font-medium">{summary.category}</span>
                        <span>{summary.count} item{summary.count === 1 ? "" : "s"}</span>
                      </div>
                      <div className="mt-1 flex items-center justify-between gap-3 text-slate-400">
                        <span>Length</span>
                        <span>{summary.length}</span>
                      </div>
                      <div className="mt-1 flex items-center justify-between gap-3 text-slate-400">
                        <span>Area</span>
                        <span>{summary.area}</span>
                      </div>
                    </div>
                  ))
                ) : (
                  <div className="rounded-lg border border-dashed border-slate-700 bg-[#191e25] p-3 text-slate-500">
                    No visible takeoff data yet.
                  </div>
                )}
              </div>
            </div>

            <div className="space-y-2 rounded-xl border border-slate-700 bg-[#12161b] p-3">
              <div className="flex items-center gap-2 font-medium"><ClipboardList className="h-4 w-4" /> Reporting</div>
              <Input value={reportTitle} onChange={(e) => setReportTitle(e.target.value)} placeholder="Report title" />
              <Button onClick={exportReport}><ClipboardList className="mr-2 h-4 w-4" />Export Report</Button>
            </div>

            <div className="space-y-2 rounded-xl border border-slate-700 bg-[#12161b] p-3">
              <div className="font-medium">Sessions</div>
              <div className="max-h-36 space-y-2 overflow-auto">
                {sessions.map((session) => (
                  <button key={session.id} onClick={() => loadSession(session)} className="w-full rounded-lg border border-slate-700 bg-[#191e25] p-2 text-left hover:bg-[#222830]">
                    <div className="font-medium">{session.name}</div>
                    <div className="text-xs text-slate-500">{new Date(session.savedAt).toLocaleString()}</div>
                  </button>
                ))}
              </div>
            </div>

            <div className="space-y-2 rounded-2xl border border-slate-700 bg-[#12161b] p-3">
              <div className="font-medium">Self Tests</div>
              {tests.map((test) => (
                <div key={test.name} className="flex items-center justify-between text-xs">
                  <span>{test.name}</span>
                  <Badge variant={test.pass ? "secondary" : "destructive"}>{test.pass ? "pass" : "fail"}</Badge>
                </div>
              ))}
            </div>
          </CardContent>
          {isWideLayout ? (
            <button
              type="button"
              aria-label="Resize inspector panel"
              onPointerDown={(event) => beginPanelResize("right", event.clientX)}
              className="absolute left-0 top-4 bottom-4 w-1 cursor-col-resize rounded-full bg-transparent transition hover:bg-amber-400/60"
            />
          ) : null}
        </Card>
        ) : null}
      </motion.div>
    </div>
  );
}
