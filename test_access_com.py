from access_com import AccessCOMManager

if __name__ == "__main__":
    mgr = AccessCOMManager()
    try:
        print("Connecting to running Access instance...")
        mgr.connect()
        print("Connected!")

        print("\nModules:")
        try:
            modules = mgr.list_modules()
            print(modules)
        except Exception as e:
            print(f"Error listing modules: {e}")

        print("\nQueries:")
        try:
            queries = mgr.list_queries()
            print(queries)
        except Exception as e:
            print(f"Error listing queries: {e}")

        print("\nForms:")
        try:
            forms = mgr.list_forms()
            print(forms)
        except Exception as e:
            print(f"Error listing forms: {e}")

        print("\nMSys Tables:")
        try:
            msys_tables = mgr.list_msys_tables()
            print(msys_tables)
        except Exception as e:
            print(f"Error listing MSys tables: {e}")

    except Exception as e:
        print(f"Failed to connect or run tests: {e}")
    finally:
        mgr.disconnect()
        print("Disconnected.")
