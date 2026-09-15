import { NextResponse } from "next/server";
import clientPromise from "@/lib/mongodb";
import { GridFSBucket, ObjectId } from "mongodb";
import { assertSuperadmin } from "@/lib/auth-server";

// PATCH /api/produk/documents/[id] -> Toggle lock
export async function PATCH(
  req: Request,
  { params }: { params: Promise<{ id: string }> }
) {
  try {
    const auth = assertSuperadmin(req);
    if (!auth.ok) {
      return NextResponse.json({ error: auth.error }, { status: auth.status });
    }

    const { id } = await params;
    if (!id || !ObjectId.isValid(id)) {
      return NextResponse.json({ error: "ID dokumen tidak valid" }, { status: 400 });
    }

    const body = await req.json();
    const { isLocked } = body;

    const client = await clientPromise;
    const db = client.db("MabelHub");
    const col = db.collection("product_documents");

    const result = await col.findOneAndUpdate(
      { _id: new ObjectId(id) },
      { $set: { isLocked: Boolean(isLocked), updatedAtDate: new Date() } },
      { returnDocument: "after" }
    );

    if (!result) {
      return NextResponse.json({ error: "Dokumen tidak ditemukan" }, { status: 404 });
    }

    return NextResponse.json({ success: true, isLocked: Boolean(isLocked) });
  } catch (error) {
    console.error("[PATCH /api/produk/documents/[id]] Error:", error);
    return NextResponse.json(
      { error: "Internal Server Error" },
      { status: 500 }
    );
  }
}

// DELETE /api/produk/documents/[id]
export async function DELETE(
  req: Request,
  { params }: { params: Promise<{ id: string }> }
) {
  try {
    const auth = assertSuperadmin(req);
    if (!auth.ok) {
      return NextResponse.json({ error: auth.error }, { status: auth.status });
    }

    const { id } = await params;
    if (!id || !ObjectId.isValid(id)) {
      return NextResponse.json({ error: "ID dokumen tidak valid" }, { status: 400 });
    }

    const client = await clientPromise;
    const db = client.db("MabelHub");
    const col = db.collection("product_documents");

    const doc = await col.findOne({ _id: new ObjectId(id) });
    if (!doc) {
      return NextResponse.json({ error: "Dokumen tidak ditemukan" }, { status: 404 });
    }

    // Delete file from GridFS bucket if fileId exists
    if (doc.fileId) {
      try {
        const bucket = new GridFSBucket(db, { bucketName: "product_files" });
        await bucket.delete(new ObjectId(doc.fileId));
      } catch (gridFsErr) {
        console.warn("GridFS delete warning:", gridFsErr);
      }
    }

    // Delete document metadata
    await col.deleteOne({ _id: new ObjectId(id) });

    // Decrement docCount in product_categories
    if (doc.categoryId) {
      await db.collection("product_categories").updateOne(
        { id: doc.categoryId },
        { $inc: { docCount: -1 } }
      );
    }

    return NextResponse.json({ success: true });
  } catch (error) {
    console.error("[DELETE /api/produk/documents/[id]] Error:", error);
    return NextResponse.json(
      { error: "Internal Server Error" },
      { status: 500 }
    );
  }
}
