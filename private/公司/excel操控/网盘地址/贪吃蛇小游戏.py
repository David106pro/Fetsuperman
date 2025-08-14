import pygame
import random
import sys

# 初始化 Pygame
pygame.init()

# 定义颜色
WHITE = (255, 255, 255)
RED = (255, 0, 0)
GREEN = (0, 255, 0)
BLACK = (0, 0, 0)
DARK_GREEN = (0, 200, 0)
YELLOW = (255, 255, 0)

# 设置游戏窗口
WINDOW_WIDTH = 800
WINDOW_HEIGHT = 600
BLOCK_SIZE = 20

# 难度设置
DIFFICULTY = {
    "简单": 10,
    "普通": 15,
    "困难": 25
}

# 创建游戏窗口
screen = pygame.display.set_mode((WINDOW_WIDTH, WINDOW_HEIGHT))
pygame.display.set_caption('贪吃蛇游戏')

# 初始化时钟
clock = pygame.time.Clock()


class Snake:
    def __init__(self):
        self.length = 1
        self.positions = [(WINDOW_WIDTH // 2, WINDOW_HEIGHT // 2)]
        self.direction = random.choice([UP, DOWN, LEFT, RIGHT])
        self.color = GREEN
        self.score = 0
        self.head_color = DARK_GREEN  # 蛇头颜色

    def get_head_position(self):
        return self.positions[0]

    def update(self):
        cur = self.get_head_position()
        x, y = self.direction
        new = ((cur[0] + (x * BLOCK_SIZE)) % WINDOW_WIDTH, (cur[1] + (y * BLOCK_SIZE)) % WINDOW_HEIGHT)
        if new in self.positions[3:]:
            return False
        else:
            self.positions.insert(0, new)
            if len(self.positions) > self.length:
                self.positions.pop()
            return True

    def reset(self):
        self.length = 1
        self.positions = [(WINDOW_WIDTH // 2, WINDOW_HEIGHT // 2)]
        self.direction = random.choice([UP, DOWN, LEFT, RIGHT])
        self.score = 0

    def draw(self, surface):
        # 绘制蛇身，使用渐变色
        for i, p in enumerate(self.positions):
            color = (
                int(GREEN[0] * (1 - i / len(self.positions))),
                int(GREEN[1]),
                int(GREEN[2] * (1 - i / len(self.positions)))
            )
            pygame.draw.rect(surface, color, (p[0], p[1], BLOCK_SIZE, BLOCK_SIZE))
            # 添加边框使蛇身更立体
            pygame.draw.rect(surface, DARK_GREEN, (p[0], p[1], BLOCK_SIZE, BLOCK_SIZE), 1)
        
        # 特别绘制蛇头
        head = self.positions[0]
        pygame.draw.rect(surface, self.head_color, (head[0], head[1], BLOCK_SIZE, BLOCK_SIZE))
        # 添加眼睛
        eye_size = 4
        if self.direction == RIGHT:
            pygame.draw.circle(surface, YELLOW, (head[0] + BLOCK_SIZE - 5, head[1] + 5), eye_size)
            pygame.draw.circle(surface, YELLOW, (head[0] + BLOCK_SIZE - 5, head[1] + BLOCK_SIZE - 5), eye_size)
        elif self.direction == LEFT:
            pygame.draw.circle(surface, YELLOW, (head[0] + 5, head[1] + 5), eye_size)
            pygame.draw.circle(surface, YELLOW, (head[0] + 5, head[1] + BLOCK_SIZE - 5), eye_size)
        elif self.direction == UP:
            pygame.draw.circle(surface, YELLOW, (head[0] + 5, head[1] + 5), eye_size)
            pygame.draw.circle(surface, YELLOW, (head[0] + BLOCK_SIZE - 5, head[1] + 5), eye_size)
        elif self.direction == DOWN:
            pygame.draw.circle(surface, YELLOW, (head[0] + 5, head[1] + BLOCK_SIZE - 5), eye_size)
            pygame.draw.circle(surface, YELLOW, (head[0] + BLOCK_SIZE - 5, head[1] + BLOCK_SIZE - 5), eye_size)


class Food:
    def __init__(self):
        self.position = (0, 0)
        self.color = RED
        self.randomize_position()

    def randomize_position(self):
        self.position = (random.randint(0, (WINDOW_WIDTH - BLOCK_SIZE) // BLOCK_SIZE) * BLOCK_SIZE,
                         random.randint(0, (WINDOW_HEIGHT - BLOCK_SIZE) // BLOCK_SIZE) * BLOCK_SIZE)

    def draw(self, surface):
        pygame.draw.rect(surface, self.color, (self.position[0], self.position[1], BLOCK_SIZE, BLOCK_SIZE))


def draw_grid(surface):
    for y in range(0, WINDOW_HEIGHT, BLOCK_SIZE):
        for x in range(0, WINDOW_WIDTH, BLOCK_SIZE):
            r = pygame.Rect(x, y, BLOCK_SIZE, BLOCK_SIZE)
            pygame.draw.rect(surface, WHITE, r, 1)


# 定义方向
UP = (0, -1)
DOWN = (0, 1)
LEFT = (-1, 0)
RIGHT = (1, 0)


def select_difficulty():
    font = pygame.font.Font(None, 48)
    difficulty_options = list(DIFFICULTY.keys())
    selected = 0
    
    while True:
        screen.fill(BLACK)
        title = font.render("选择难度", True, WHITE)
        screen.blit(title, (WINDOW_WIDTH//2 - title.get_width()//2, 100))
        
        for i, diff in enumerate(difficulty_options):
            color = GREEN if i == selected else WHITE
            text = font.render(diff, True, color)
            screen.blit(text, (WINDOW_WIDTH//2 - text.get_width()//2, 250 + i * 60))
        
        pygame.display.update()
        
        for event in pygame.event.get():
            if event.type == pygame.QUIT:
                pygame.quit()
                sys.exit()
            if event.type == pygame.KEYDOWN:
                if event.key == pygame.K_UP:
                    selected = (selected - 1) % len(difficulty_options)
                elif event.key == pygame.K_DOWN:
                    selected = (selected + 1) % len(difficulty_options)
                elif event.key == pygame.K_RETURN:
                    return difficulty_options[selected]


def main():
    # 选择难度
    difficulty = select_difficulty()
    GAME_SPEED = DIFFICULTY[difficulty]
    
    snake = Snake()
    food = Food()
    font = pygame.font.Font(None, 36)

    while True:
        for event in pygame.event.get():
            if event.type == pygame.QUIT:
                pygame.quit()
                sys.exit()
            elif event.type == pygame.KEYDOWN:
                if event.key == pygame.K_UP and snake.direction != DOWN:
                    snake.direction = UP
                elif event.key == pygame.K_DOWN and snake.direction != UP:
                    snake.direction = DOWN
                elif event.key == pygame.K_LEFT and snake.direction != RIGHT:
                    snake.direction = LEFT
                elif event.key == pygame.K_RIGHT and snake.direction != LEFT:
                    snake.direction = RIGHT

        # 更新蛇的位置
        if not snake.update():
            snake.reset()
            food.randomize_position()

        # 检查是否吃到食物
        if snake.get_head_position() == food.position:
            snake.length += 1
            snake.score += 1
            food.randomize_position()

        # 绘制游戏界面
        screen.fill(BLACK)
        draw_grid(screen)
        snake.draw(screen)
        food.draw(screen)

        # 显示分数和当前难度
        score_text = font.render(f'得分: {snake.score}', True, WHITE)
        diff_text = font.render(f'难度: {difficulty}', True, WHITE)
        screen.blit(score_text, (5, 5))
        screen.blit(diff_text, (5, 35))

        pygame.display.update()
        clock.tick(GAME_SPEED)


if __name__ == '__main__':
    main()