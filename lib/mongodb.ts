import mongoose from "mongoose";

type MongooseCache = {
  connection: typeof mongoose | null;
  promise: Promise<typeof mongoose> | null;
};

declare global {
  var mongooseCache: MongooseCache | undefined;
}

const cached = global.mongooseCache ?? { connection: null, promise: null };
global.mongooseCache = cached;

export default async function connectToDatabase() {
  if (cached.connection) return cached.connection;

  if (!cached.promise) {
    const mongoUri = process.env.MONGODB_URI;
    if (!mongoUri) {
      throw new Error("Please define the MONGODB_URI environment variable.");
    }

    const isSrvUri = mongoUri.startsWith("mongodb+srv://");
    cached.promise = mongoose
      .connect(mongoUri, {
        serverSelectionTimeoutMS: 15_000,
        connectTimeoutMS: 10_000,
        family: 4,
      })
      .catch((error: unknown) => {
        cached.promise = null;

        const errorMessage = error instanceof Error ? error.message : String(error);
        if (
          isSrvUri &&
          /querySrv|ECONNREFUSED|ENOTFOUND|EAI_AGAIN/i.test(errorMessage)
        ) {
          console.error(
            "MongoDB Atlas SRV DNS lookup failed. Check your internet connection, Atlas cluster status, and Network Access IP allowlist. If your DNS provider blocks SRV records, use the standard mongodb:// connection string from Atlas.",
            error,
          );
        } else {
          console.error("MongoDB connection failed. Check the URI, Atlas cluster, and network access.", error);
        }

        throw error;
      });
  }

  cached.connection = await cached.promise;
  return cached.connection;
}
