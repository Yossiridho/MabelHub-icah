import { NextResponse } from "next/server";
import clientPromise from "@/lib/mongodb";
import { getSession } from "@/lib/auth-server";

export async function GET(
  req: Request,
  { params }: { params: Promise<{ id: string }> },
) {
  try {
    const { id } = await params;

    if (!id) {
      return NextResponse.json(
        { error: "Category ID is required" },
        { status: 400 },
      );
    }

    const client = await clientPromise;
    const db = client.db("MabelHub");
    const col = db.collection("product_categories");

    const category = await col.findOne({ id });

    if (!category) {
      return NextResponse.json(
        { error: "Category not found" },
        { status: 404 },
      );
    }

    return NextResponse.json({
      id: category.id,
      name: category.name,
      description: category.description || "",
      isLocked: category.isLocked || false,
    });
  } catch (error) {
    console.error("[GET /api/produk/categories/[id]] Error:", error);
    return NextResponse.json(
      { error: "Internal Server Error" },
      { status: 500 },
    );
  }
}

export async function PATCH(
  req: Request,
  { params }: { params: Promise<{ id: string }> },
) {
  try {
    const { id } = await params;

    if (!id) {
      return NextResponse.json(
        { error: "Category ID is required" },
        { status: 400 },
      );
    }

    const body = await req.json();
    const { isLocked } = body;

    if (typeof isLocked !== "boolean") {
      return NextResponse.json(
        { error: "isLocked must be a boolean" },
        { status: 400 },
      );
    }

    const client = await clientPromise;
    const db = client.db("MabelHub");
    const col = db.collection("product_categories");

    const result = await col.updateOne({ id }, { $set: { isLocked } });

    if (result.matchedCount === 0) {
      return NextResponse.json(
        { error: "Category not found" },
        { status: 404 },
      );
    }

    return NextResponse.json({ success: true, isLocked });
  } catch (error) {
    console.error("[PATCH /api/produk/categories/[id]] Error:", error);
    return NextResponse.json(
      { error: "Internal Server Error" },
      { status: 500 },
    );
  }
}

export async function DELETE(
  req: Request,
  { params }: { params: Promise<{ id: string }> },
) {
  try {
    const session = getSession(req);

    if (!session || session.role !== "SUPERADMIN") {
      return NextResponse.json(
        { error: "Unauthorized. Only SUPERADMIN can delete categories." },
        { status: 403 },
      );
    }

    const { id } = await params;

    if (!id) {
      return NextResponse.json(
        { error: "Category ID is required" },
        { status: 400 },
      );
    }

    const client = await clientPromise;
    const db = client.db("MabelHub");
    const catCol = db.collection("product_categories");
    const docCol = db.collection("product_documents");

    // Check if category exists
    const category = await catCol.findOne({ id });
    if (!category) {
      return NextResponse.json(
        { error: "Category not found" },
        { status: 404 },
      );
    }

    // Delete all documents in this category
    await docCol.deleteMany({ categoryId: id });

    // Delete the category itself
    await catCol.deleteOne({ id });

    return NextResponse.json({ success: true, deletedId: id });
  } catch (error) {
    console.error("[DELETE /api/produk/categories/[id]] Error:", error);
    return NextResponse.json(
      { error: "Internal Server Error" },
      { status: 500 },
    );
  }
}
