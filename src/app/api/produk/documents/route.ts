import { NextResponse } from "next/server";
import clientPromise from "@/lib/mongodb";
import { GridFSBucket, ObjectId } from "mongodb";
import { assertLoggedIn, assertSuperadmin } from "@/lib/auth-server";

function formatFileSize(bytes: number): string {
  if (!bytes || bytes === 0) return "0 B";
  const k = 1024;
  const sizes = ["B", "KB", "MB", "GB"];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + " " + sizes[i];
}

function getFileType(fileName: string): string {
  const parts = fileName.split(".");
  if (parts.length > 1) {
    return parts.pop()!.toUpperCase();
  }
  return "FILE";
}

// GET /api/produk/documents?kategori=...
export async function GET(req: Request) {
  try {
    const auth = assertLoggedIn(req);
    if (!auth.ok) {
      return NextResponse.json({ error: auth.error }, { status: auth.status });
    }

    const { searchParams } = new URL(req.url);
    const categoryId = searchParams.get("kategori");

    if (!categoryId) {
      return NextResponse.json(
        { error: "Parameter 'kategori' is required" },
        { status: 400 }
      );
    }

    const client = await clientPromise;
    const db = client.db("MabelHub");
    const col = db.collection("product_documents");

    const isSalesOrLeader =
      auth.session.role === "SALES" || auth.session.role === "LEADER";

    const query: any = { categoryId };
    if (isSalesOrLeader) {
      query.isLocked = { $ne: true };
    }

    const rawDocs = await col
      .find(query)
      .sort({ createdAt: -1 })
      .toArray();

    const formatted = rawDocs.map((d: any) => ({
      id: d._id.toString(),
      name: d.name,
      type: d.type || getFileType(d.name),
      size: d.size || formatFileSize(d.sizeBytes || 0),
      updatedAt: d.updatedAt || "",
      url: `/api/produk/documents/${d._id.toString()}/file`,
      isLocked: Boolean(d.isLocked),
    }));

    return NextResponse.json(formatted);
  } catch (error) {
    console.error("[GET /api/produk/documents] Error:", error);
    return NextResponse.json(
      { error: "Internal Server Error" },
      { status: 500 }
    );
  }
}

// POST /api/produk/documents (FormData: file, categoryId)
export async function POST(req: Request) {
  try {
    const auth = assertSuperadmin(req);
    if (!auth.ok) {
      return NextResponse.json({ error: auth.error }, { status: auth.status });
    }

    const formData: any = await req.formData();
    const file = formData.get("file") as File | null;
    const categoryId = formData.get("categoryId") as string | null;

    if (!file || !categoryId) {
      return NextResponse.json(
        { error: "File dan categoryId wajib disertakan" },
        { status: 400 }
      );
    }

    const client = await clientPromise;
    const db = client.db("MabelHub");

    // Convert file to buffer
    const arrayBuffer = await file.arrayBuffer();
    const buffer = Buffer.from(arrayBuffer);

    // Save file stream into GridFS
    const bucket = new GridFSBucket(db, { bucketName: "product_files" });
    const uploadStream = bucket.openUploadStream(file.name, {
      contentType: file.type || "application/octet-stream",
      metadata: {
        categoryId,
        originalName: file.name,
        sizeBytes: file.size,
        uploadedBy: auth.session?.username || auth.session?.userId || "unknown",
        createdAt: new Date(),
      },
    });

    await new Promise<void>((resolve, reject) => {
      uploadStream.end(buffer, (err?: Error | null) => {
        if (err) reject(err);
        else resolve();
      });
    });

    const fileId = uploadStream.id;
    const fileType = getFileType(file.name);
    const formattedSize = formatFileSize(file.size);
    const formattedDate = new Date().toLocaleDateString("id-ID", {
      day: "numeric",
      month: "short",
      year: "numeric",
    });

    const newDoc = {
      categoryId,
      name: file.name,
      type: fileType,
      size: formattedSize,
      sizeBytes: file.size,
      updatedAt: formattedDate,
      fileId,
      isLocked: false,
      uploadedBy: auth.session?.username || auth.session?.userId || "unknown",
      createdAt: new Date(),
    };

    const result = await db.collection("product_documents").insertOne(newDoc);

    // Update docCount in product_categories
    await db.collection("product_categories").updateOne(
      { id: categoryId },
      { $inc: { docCount: 1 } }
    );

    const createdDoc = {
      id: result.insertedId.toString(),
      name: newDoc.name,
      type: newDoc.type,
      size: newDoc.size,
      updatedAt: newDoc.updatedAt,
      url: `/api/produk/documents/${result.insertedId.toString()}/file`,
      isLocked: newDoc.isLocked,
    };

    return NextResponse.json(
      { success: true, document: createdDoc },
      { status: 201 }
    );
  } catch (error) {
    console.error("[POST /api/produk/documents] Error:", error);
    return NextResponse.json(
      { error: "Gagal mengunggah dokumen" },
      { status: 500 }
    );
  }
}
