import { NextResponse } from "next/server";
import clientPromise from "@/lib/mongodb";
import { GridFSBucket, ObjectId } from "mongodb";
import { assertLoggedIn } from "@/lib/auth-server";
import { Readable } from "stream";

const MIME_MAP: Record<string, string> = {
  pdf: "application/pdf",
  png: "image/png",
  jpg: "image/jpeg",
  jpeg: "image/jpeg",
  gif: "image/gif",
  webp: "image/webp",
  svg: "image/svg+xml",
  xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  xls: "application/vnd.ms-excel",
  docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  doc: "application/msword",
  pptx: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  ppt: "application/vnd.ms-powerpoint",
  txt: "text/plain",
  csv: "text/csv",
};

export async function GET(
  req: Request,
  { params }: { params: Promise<{ id: string }> }
) {
  try {
    const auth = assertLoggedIn(req);
    if (!auth.ok) {
      return NextResponse.json({ error: auth.error }, { status: auth.status });
    }

    const { id } = await params;
    if (!id || !ObjectId.isValid(id)) {
      return NextResponse.json({ error: "ID dokumen tidak valid" }, { status: 400 });
    }

    const client = await clientPromise;
    const db = client.db("MabelHub");
    const doc = await db.collection("product_documents").findOne({ _id: new ObjectId(id) });

    if (!doc || !doc.fileId) {
      return NextResponse.json({ error: "File tidak ditemukan" }, { status: 404 });
    }

    // Check lock restriction
    const isSalesOrLeader =
      auth.session.role === "SALES" || auth.session.role === "LEADER";

    if (doc.isLocked && isSalesOrLeader) {
      return NextResponse.json(
        { error: "Dokumen ini terkunci dan tidak dapat diakses" },
        { status: 403 }
      );
    }

    const isSuperAdmin = auth.session.role === "SUPERADMIN";
    const { searchParams } = new URL(req.url);
    const isDownload = searchParams.get("download") === "true" || searchParams.get("download") === "1";

    if (doc.isLocked && !isSuperAdmin && isDownload) {
      return NextResponse.json(
        { error: "Dokumen ini terkunci dan tidak dapat diunduh" },
        { status: 403 }
      );
    }

    const bucket = new GridFSBucket(db, { bucketName: "product_files" });
    const fileId = new ObjectId(doc.fileId);

    const fileRecords = await db
      .collection("product_files.files")
      .find({ _id: fileId })
      .toArray();

    if (!fileRecords || fileRecords.length === 0) {
      return NextResponse.json({ error: "Data file fisik tidak ditemukan" }, { status: 404 });
    }

    const fileRecord = fileRecords[0];
    const ext = (doc.name.split(".").pop() || "").toLowerCase();
    const contentType =
      fileRecord.contentType || MIME_MAP[ext] || "application/octet-stream";

    const downloadStream = bucket.openDownloadStream(fileId);
    const webStream = Readable.toWeb(downloadStream as any);

    const filename = doc.name || fileRecord.filename || "dokumen";
    const dispositionType = isDownload ? "attachment" : "inline";

    return new NextResponse(webStream as any, {
      headers: {
        "Content-Type": contentType,
        "Content-Length": String(fileRecord.length),
        "Content-Disposition": `${dispositionType}; filename="${encodeURIComponent(
          filename
        )}"`,
        "Cache-Control": "public, max-age=3600",
      },
    });
  } catch (error) {
    console.error("[GET /api/produk/documents/[id]/file] Error:", error);
    return NextResponse.json(
      { error: "Internal Server Error" },
      { status: 500 }
    );
  }
}
