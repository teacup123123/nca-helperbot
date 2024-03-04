import dataclasses


def transition_from(state0):
    def decorator(func):
        def decorated(self, *args, **kwargs):
            return func(self, *args, **kwargs)

        decorated._src_state = state0
        return decorated

    return decorator


@dataclasses.dataclass
class bot:
    state: str = 'init'

    def handle(self, *args, **kwargs):
        cls = type(self)
        dirs = dir(cls)
        if 'handles' not in dirs:
            self.handles = {}
            for attr in filter(lambda x: not x.startswith('__'), dirs):
                if attr in cls.__dict__ and type(clsmethod := cls.__dict__[attr]) == type(lambda x: x):
                    if '_src_state' in clsmethod.__dict__:
                        self.handles[clsmethod.__dict__['_src_state']] = self.__getattribute__(attr)
        res = self.handles[self.state](*args, **kwargs)
        self.state, _reply = res
        return res


if __name__ == '__main__':
    class mybot(bot):

        @transition_from('init')
        def test_transition(self, msg):
            return 'next', 'going to next'


    b = mybot()
    b.handle('sdf')
    print(b.state)
