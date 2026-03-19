from manim import *
import numpy as np

class BinaryCounter(Scene):
    def construct(self):
        # Configuration
        num_bits = 7
        total_duration = 45  # seconds (within 30-60 sec range)
        hold_time = 0.3      # Time to hold each number
        transition_time = 0.15  # Time for bit-flip animation

        # Create binary register and decimal number
        binary_group = VGroup()
        decimal_number = Text("", font="Courier New", font_size=72, color=WHITE)
        
        # Initial state (0)
        current_value = 0
        binary_str = format(current_value, '07b')
        binary_group = self.create_binary_group(binary_str, num_bits)
        decimal_number.next_to(binary_group, DOWN, buff=0.3)
        
        # Position in center of screen
        binary_group.move_to(UP * 0.5)
        
        # Add to scene
        self.add(binary_group, decimal_number)
        self.wait(hold_time)
        
        # Animate counting from 0 to 127
        for next_value in range(1, 128):
            # Calculate bits that change from current to next
            current_bin = format(current_value, '07b')
            next_bin = format(next_value, '07b')
            flipping_bits = []
            
            for i in range(num_bits):
                if current_bin[i] == '0' and next_bin[i] == '1':
                    flipping_bits.append(i)
            
            # Animate transition to next value
            new_binary_group = self.create_binary_group(next_bin, num_bits)
            new_binary_group.move_to(binary_group)
            transition_anims = []
            
            for i, bit in enumerate(next_bin):
                if i in flipping_bits:
                    # Flash animation for 0->1 bits
                    orig_bit = binary_group[i]
                    flash = orig_bit.copy()
                    flash.set_color(GREEN)
                    flash.animate.set_color(WHITE).set_run_time(transition_time)
                    transition_anims.append(
                        AnimationGroup(
                            orig_bit.animate.set_color(GREEN),
                            run_time=transition_time/3
                        ) + 
                        flash.animate.set_color(WHITE).set_run_time(transition_time/3)
                    )
            
            # Update binary group and decimal number
            transition_anims.append(
                binary_group.animate.replace(new_binary_group)
            )
            transition_anims.append(
                decimal_number.animate.set_text(str(next_value))
            )
            
            # Play all animations simultaneously
            self.play(*transition_anims, run_time=transition_time)
            self.wait(hold_time)
            current_value = next_value
        
        # Hold final frame (127)
        self.wait(2)

    def create_binary_group(self, binary_str, num_bits):
        """Create VGroup with individual bit texts"""
        group = VGroup()
        spacing = 0.8
        start_x = -3 * spacing
        
        for i, bit in enumerate(binary_str):
            bit_text = Text(bit, font="Courier New", font_size=80, color=WHITE)
            bit_text.move_to(LEFT * (num_bits - i - 1) * spacing)
            group.add(bit_text)
        
        return group

# Render the video
if __name__ == "__main__":
    config.background_color = BLACK
    config.frame_rate = 30
    config.pixel_height = 1080
    config.pixel_width = 1920
    
    scene = BinaryCounter()
    scene.render()