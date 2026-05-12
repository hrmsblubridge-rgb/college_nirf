"""
Migrate local MongoDB data to live MongoDB Atlas.
Usage: python3 migrate_to_atlas.py

Reads LIVE_MONGO_URL and LIVE_DB_NAME from backend/.env
Copies all collections from local DB to Atlas.
"""
import pymongo
import os
import sys
from pathlib import Path
from dotenv import load_dotenv

load_dotenv(Path(__file__).parent / 'backend' / '.env')

LOCAL_URL = os.environ.get('MONGO_URL', 'mongodb://localhost:27017')
LOCAL_DB = os.environ.get('DB_NAME', 'test_database')
LIVE_URL = os.environ.get('LIVE_MONGO_URL')
LIVE_DB = os.environ.get('LIVE_DB_NAME', 'nirf')

if not LIVE_URL:
    print("ERROR: LIVE_MONGO_URL not set in backend/.env")
    sys.exit(1)

print(f"Source: {LOCAL_URL} / {LOCAL_DB}")
print(f"Target: {LIVE_URL[:50]}... / {LIVE_DB}")
print()

# Connect to local
try:
    local_client = pymongo.MongoClient(LOCAL_URL, serverSelectionTimeoutMS=5000)
    local_client.admin.command('ping')
    local_db = local_client[LOCAL_DB]
    print("Local MongoDB: CONNECTED")
except Exception as e:
    print(f"Local MongoDB FAILED: {e}")
    sys.exit(1)

# Connect to Atlas
try:
    live_client = pymongo.MongoClient(LIVE_URL, serverSelectionTimeoutMS=15000)
    live_client.admin.command('ping')
    live_db = live_client[LIVE_DB]
    print("Atlas MongoDB: CONNECTED")
except Exception as e:
    print(f"Atlas MongoDB FAILED: {e}")
    print("\nPlease verify:")
    print("  1. Username/password in LIVE_MONGO_URL")
    print("  2. Network Access: add 0.0.0.0/0 in Atlas")
    sys.exit(1)

# Migrate collections
collections = local_db.list_collection_names()
print(f"\nCollections to migrate: {collections}")

for coll_name in collections:
    local_coll = local_db[coll_name]
    live_coll = live_db[coll_name]
    count = local_coll.count_documents({})

    if count == 0:
        print(f"  {coll_name}: SKIPPED (empty)")
        continue

    # Drop existing in Atlas
    live_coll.drop()

    # Copy in batches
    batch_size = 5000
    total = 0
    docs = list(local_coll.find({}))
    # Remove _id to let Atlas generate new ones
    for doc in docs:
        doc.pop('_id', None)

    for i in range(0, len(docs), batch_size):
        batch = docs[i:i+batch_size]
        live_coll.insert_many(batch)
        total += len(batch)

    verify = live_coll.count_documents({})
    status = "OK" if verify == count else f"MISMATCH (expected {count})"
    print(f"  {coll_name}: {verify} documents migrated - {status}")

    # Copy indexes
    for idx in local_coll.list_indexes():
        if idx['name'] == '_id_':
            continue
        try:
            live_coll.create_index(list(idx['key'].items()), name=idx['name'])
        except Exception:
            pass

print("\nMigration complete!")
print("\nTo switch to Atlas, update backend/.env:")
print(f'  MONGO_URL="{LIVE_URL}"')
print(f'  DB_NAME="{LIVE_DB}"')

local_client.close()
live_client.close()
