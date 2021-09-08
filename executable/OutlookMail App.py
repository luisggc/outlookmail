if __name__ == '__main__':
    try:
        import outlookmail as om
        filename = "settings.txt"
        om.API.read_from_txt(filename, success_message=True)
    except BaseException:
        import sys
        print(sys.exc_info()[0])
        import traceback
        print(traceback.format_exc())
    finally:
        print("Press Enter to exit ...")
        text = input()