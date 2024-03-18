import inspect
import traceback

functions = []


def interaction(fun):
    functions.append(fun)
    return fun


def start_interaction(input=input):
    while True:
        try:
            print('\nOptions are as follows:')
            for fi, fn in enumerate(functions):
                if len(functions) <= 10:
                    print(f'[{fi:01}] {fn.__name__}{inspect.signature(fn)}')
                else:
                    print(f'[{fi:02}] {fn.__name__}{inspect.signature(fn)}')
            print('.....what to choose?', end='')
            f = functions[int(input())]
            signature = inspect.signature(f)
            print(f'reminder, signature are as follows: {f.__name__}{signature}')
            args = []
            for param in signature.parameters:
                print('.....' + param + '=?', end='')
                args.append(input())
            f(*args)
        except KeyboardInterrupt:
            try:
                input('** Ctrl+C again to exit. enter to restart **')
            except KeyboardInterrupt:
                break
            continue
        except Exception as e:
            print(e)
            print(traceback.format_exc())
            print('continuing despite error...')
            continue


if __name__ == '__main__':
    start_interaction()
