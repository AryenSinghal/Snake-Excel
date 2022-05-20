import xlwings as xl
from random import randint
from time import sleep
from sys import exit

# COLORS
BOARD_COLOR = (226,227,223)
CONTROL_COLOR = (46,50,51)
BUTTON_COLOR = (81,91,94)
TEXT_COLOR = (255,255,255)
APPLE_COLOR = (0,255,100)
HEAD_COLOR = (255,0,0)
BODY_COLOR = (200,0,0)

class Snake:
    def __init__(self, speed, width, height):
        # book setup
        self.book = xl.Book()
        self.sheet = self.book.sheets[0]
        self.sheet.name = 'Snake'

        # board setup
        self.speed = speed
        self.width = width
        self.height = height

        # base colors
        self.sheet.range((2,2),(self.height+1,self.width+1)).color = BOARD_COLOR
        self.sheet.range((self.height+2,2),(self.height+6,self.width+1)).color = CONTROL_COLOR

        # buttons
        self.exit_cell = (self.height+6,self.width+1)
        self.sheet.range(self.exit_cell).value = 'quit'

        self.left_cell = (self.height+4,3)
        self.sheet.range(self.left_cell).value = 'left'

        self.right_cell = (self.height+4,5)
        self.sheet.range(self.right_cell).value = 'right'

        self.up_cell = (self.height+3,4)
        self.sheet.range(self.up_cell).value = 'up'

        self.down_cell = (self.height+5,4)
        self.sheet.range(self.down_cell).value = 'down'

        for cell in (self.exit_cell, self.left_cell, self.right_cell, self.up_cell, self.down_cell):
            self.sheet.range(cell).color = BUTTON_COLOR
            self.sheet.range(cell).font.color = TEXT_COLOR
        
        # cell dimensions
        self.sheet.range((2,2),(self.height+6,2)).row_height = 40

        # snake and apple setup
        self.body = [(self.height//2,5), (self.height//2,4), (self.height//2,3)]
        self.direction = (0,1)
        self.eaten = False
        self.create_apple()
        self.display_snake()    
    
    def display_snake(self):
        self.sheet[self.body[0]].color = HEAD_COLOR
        for cell in self.body[1:]:
            self.sheet[cell].color = BODY_COLOR
    
    def update_snake(self):
        if self.eaten:
            new_body = self.body[:]
            self.eaten = False
        else:
            new_body = self.body[:-1]
            self.sheet[self.body[-1]].color = BOARD_COLOR
        
        new_head = self.add_direction(new_body[0], self.direction)
        new_body.insert(0, new_head)
        self.body = new_body

    def add_direction(self, cell, direction):
        row = cell[0] + direction[0]
        col = cell[1] + direction[1]
        return (row,col)

    def create_apple(self):
        row = randint(1, self.height)
        col = randint(1, self.width)

        while (row,col) in self.body:
            row = randint(1, self.height)
            col = randint(1, self.width)

        self.apple_pos = (row,col)
        self.sheet[self.apple_pos].color = APPLE_COLOR
    
    def check_apple_collision(self):
        if self.body[0] == self.apple_pos:
            self.eaten = True
            self.create_apple()

    def input(self):
        selected_cell = self.book.selection.address.replace('$','')
        if self.sheet[selected_cell] == self.sheet.range(self.left_cell):
            self.direction = (0,-1)
        elif self.sheet[selected_cell] == self.sheet.range(self.right_cell):
            self.direction = (0,1)
        elif self.sheet[selected_cell] == self.sheet.range(self.up_cell):
            self.direction = (-1,0)
        elif self.sheet[selected_cell] == self.sheet.range(self.down_cell):
            self.direction = (1,0)
    
    def exit_game(self):
        selected_cell = self.book.selection.address.replace('$','')
        if self.sheet[selected_cell] == self.sheet.range(self.exit_cell):
            self.book.close()
            exit()

    def check_death(self):
        head = self.body[0]
        body = self.body[1:]
        
        if (head in body) or (head[1] <= 0) or (head[1] > self.width) or (head[0] <= 0) or (head[0] > self.height):
            self.book.close()
            exit()

    def run(self):
        while True:
            self.exit_game()
            sleep(self.speed)
            self.input()
            self.update_snake()
            self.check_apple_collision()
            self.check_death()
            self.display_snake()

snake = Snake(0.3, 12, 8)
snake.run()