from textual.app import App
from textual import events

class EventApp(App):
    colors = ['white', 'maroon', 'red', 'orange', 'yellow']

    def on_mount(self) -> None:
        self.screen.styles.background = 'darkblue'
    def on_key(self, event : events.Key) -> None:
        if event.key.isdecimal():
            self.screen.styles.background = self.colors[int(event.key)]
            

if __name__  ==  '__main__':
    app = EventApp()
    app.run()

