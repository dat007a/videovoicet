import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import os
import requests
import ffmpeg
from PIL import Image
import asyncio
import uuid
import json
import shutil
import random
from threading import Thread

class VideoCreatorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Video Creator from Images and Voice")
        self.api_keys = []
        self.setup_ui()
        self.output_dir = "temp_output"
        os.makedirs(self.output_dir, exist_ok=True)

    def setup_ui(self):
        tk.Label(self.root, text="ElevenLabs API Keys:").grid(row=0, column=0, padx=5, pady=5)
        tk.Button(self.root, text="Import API Keys", command=self.import_api_keys).grid(row=0, column=1, padx=5, pady=5)

        tk.Label(self.root, text="Voice ID:").grid(row=1, column=0, padx=5, pady=5)
        self.voice_id_entry = tk.Entry(self.root)
        self.voice_id_entry.grid(row=1, column=1, padx=5, pady=5)

        tk.Label(self.root, text="Stability:").grid(row=2, column=0, padx=5, pady=5)
        self.stability_scale = tk.Scale(self.root, from_=0.0, to=1.0, resolution=0.1, orient=tk.HORIZONTAL)
        self.stability_scale.set(0.5)
        self.stability_scale.grid(row=2, column=1, padx=5, pady=5)

        tk.Label(self.root, text="Similarity:").grid(row=3, column=0, padx=5, pady=5)
        self.similarity_scale = tk.Scale(self.root, from_=0.0, to=1.0, resolution=0.1, orient=tk.HORIZONTAL)
        self.similarity_scale.set(0.5)
        self.similarity_scale.grid(row=3, column=1, padx=5, pady=5)

        tk.Label(self.root, text="Speed:").grid(row=4, column=0, padx=5, pady=5)
        self.speed_scale = tk.Scale(self.root, from_=0.5, to=2.0, resolution=0.1, orient=tk.HORIZONTAL)
        self.speed_scale.set(1.0)
        self.speed_scale.grid(row=4, column=1, padx=5, pady=5)

        tk.Button(self.root, text="Check Remaining Minutes", command=self.check_remaining_minutes).grid(row=5, column=0, columnspan=2, padx=5, pady=5)

        tk.Button(self.root, text="Import Excel File", command=self.import_excel).grid(row=6, column=0, columnspan=2, padx=5, pady=5)
        tk.Button(self.root, text="Import Image Folder", command=self.import_image_folder).grid(row=7, column=0, columnspan=2, padx=5, pady=5)

        tk.Button(self.root, text="Export Output", command=self.export_output_location).grid(row=8, column=0, padx=5, pady=5)
        tk.Button(self.root, text="Create Video", command=self.start_processing).grid(row=8, column=1, padx=5, pady=5)
        self.export_dir = None

        self.progress = ttk.Progressbar(self.root, length=200, mode='determinate')
        self.progress.grid(row=9, column=0, columnspan=2, padx=5, pady=5)
        self.log_text = tk.Text(self.root, height=10, width=50)
        self.log_text.grid(row=10, column=0, columnspan=2, padx=5, pady=5)

    def log(self, message):
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.root.update()

    def import_api_keys(self):
        file_path = filedialog.askopenfilename(filetypes=[("Text Files", "*.txt")])
        if file_path:
            with open(file_path, 'r') as f:
                self.api_keys = [line.strip() for line in f if line.strip()]
            self.log(f"Imported {len(self.api_keys)} API keys.")

    def check_remaining_minutes(self):
        total_chars = 0
        for api_key in self.api_keys:
            try:
                headers = {"xi-api-key": api_key}
                response = requests.get("https://api.elevenlabs.io/v1/user/subscription", headers=headers)
                if response.status_code == 200:
                    data = response.json()
                    total_chars += data.get("character_limit", 0) - data.get("character_count", 0)
                else:
                    self.log(f"Error checking API key: {response.text}")
            except Exception as e:
                self.log(f"Error checking API key: {str(e)}")
        minutes = total_chars / 10000
        messagebox.showinfo("Remaining Minutes", f"Estimated remaining minutes: {minutes:.2f}")

    def import_excel(self):
        self.excel_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
        if self.excel_path:
            self.log(f"Imported Excel file: {self.excel_path}")

    def import_image_folder(self):
        self.image_folder = filedialog.askdirectory()
        if self.image_folder:
            self.log(f"Imported image folder: {self.image_folder}")

    def export_output_location(self):
        self.export_dir = filedialog.askdirectory()
        if self.export_dir:
            self.log(f"Selected export folder: {self.export_dir}")

    async def create_voice(self, text, output_file, api_key):
        headers = {"xi-api-key": api_key}
        payload = {
            "text": text,
            "voice_settings": {
                "stability": self.stability_scale.get(),
                "similarity_boost": self.similarity_scale.get(),
                "speed": self.speed_scale.get()
            }
        }
        voice_id = self.voice_id_entry.get() or "default_voice_id"
        try:
            response = requests.post(f"https://api.elevenlabs.io/v1/text-to-speech/{voice_id}", json=payload, headers=headers)
            if response.status_code == 200:
                with open(output_file, 'wb') as f:
                    f.write(response.content)
                self.log(f"Created voice file: {output_file}")
                return True
            else:
                self.log(f"Error creating voice: {response.text}")
                return False
        except Exception as e:
            self.log(f"Error creating voice: {str(e)}")
            return False

    def create_segment_video(self, image_path, audio_path, output_path, duration):
        try:
            # Process and save the image
            img = Image.open(image_path)
            if img.mode == 'RGBA':
                img = img.convert('RGB')
            img = img.resize((1920, 1080), Image.Resampling.LANCZOS)
            temp_img = os.path.join(self.output_dir, f"temp_{uuid.uuid4().hex}.jpg")
            img.save(temp_img, format='JPEG')

            # Enhanced zoom effect with detailed logging
            zoom_speed = min(0.01, 0.1 / max(duration, 1))  # Ensure noticeable zoom, avoid division by zero
            max_zoom = 2.0  # Increased to 2.0x for more pronounced effect
            total_frames = int(duration * 30)  # Assuming 30 fps
            final_zoom = 1.0 + (zoom_speed * total_frames)  # Calculate expected final zoom
            self.log(f"Creating video: duration={duration}s, zoom_speed={zoom_speed}, max_zoom={max_zoom}, total_frames={total_frames}, expected_final_zoom={final_zoom:.2f}")
            
            stream = ffmpeg.input(temp_img, loop=1, t=duration)
            stream = ffmpeg.filter(
                stream, 'zoompan',
                z=f'if(gte(zoom,{max_zoom}),{max_zoom},zoom+{zoom_speed})',  # Linear zoom with cap
                x='iw/2-(iw/zoom/2)',  # Center x
                y='ih/2-(ih/zoom/2)',  # Center y
                d=1,                    # Update every frame for smooth zoom
                s='1920x1080',         # Output size
                fps=30                  # Frame rate for smooth output
            )

            # Combine video and audio streams
            audio_stream = ffmpeg.input(audio_path)
            stream = ffmpeg.output(
                stream,
                audio_stream,
                output_path,
                vcodec='libx264',
                acodec='copy',
                t=duration,
                pix_fmt='yuv420p'  # Ensure compatibility with most players
            )

            # Run FFmpeg with detailed error capture
            try:
                ffmpeg.run(stream, capture_stdout=True, capture_stderr=True, quiet=True, overwrite_output=True)
                self.log(f"Created segment video: {output_path}")
                return True
            except ffmpeg.Error as e:
                self.log(f"FFmpeg error: {e.stderr.decode()}")
                return False
        except Exception as e:
            self.log(f"Error creating segment video: {str(e)}")
            return False
        finally:
            if os.path.exists(temp_img):
                os.remove(temp_img)

    def combine_videos(self, video_files, output_file):
        try:
            valid_video_files = []
            for video_file in video_files:
                if os.path.exists(video_file):
                    try:
                        ffmpeg.probe(video_file)
                        valid_video_files.append(video_file)
                        self.log(f"Validated video file: {video_file}")
                    except ffmpeg.Error as e:
                        self.log(f"Invalid video file {video_file}: {e.stderr.decode()}")
                else:
                    self.log(f"Video file not found: {video_file}")

            if not valid_video_files:
                self.log("No valid video files to combine.")
                return

            temp_file_list = os.path.join(self.output_dir, f"file_list_{uuid.uuid4().hex}.txt")
            with open(temp_file_list, 'w') as f:
                for video_file in valid_video_files:
                    normalized_path = os.path.abspath(video_file).replace('\\', '/')
                    f.write(f"file '{normalized_path}'\n")

            stream = ffmpeg.input(temp_file_list, format='concat', safe=0)
            stream = ffmpeg.output(stream, output_file, c='copy')
            ffmpeg.run(stream, overwrite_output=True, capture_stderr=True, quiet=True)
            os.remove(temp_file_list)
            self.log(f"Created final video: {output_file}")
        except ffmpeg.Error as e:
            self.log(f"FFMPEG error in combine_videos: {e.stderr.decode()}")
        except Exception as e:
            self.log(f"Error combining videos: {str(e)}")

    def start_processing(self):
        if not hasattr(self, 'excel_path') or not hasattr(self, 'image_folder'):
            messagebox.showerror("Error", "Please import Excel file and image folder.")
            return
        if not self.api_keys:
            messagebox.showerror("Error", "Please import API keys.")
            return
        Thread(target=self.process).start()

    def process(self):
        try:
            df = pd.read_excel(self.excel_path)
            if "image name" not in df.columns or "text to voice" not in df.columns:
                self.log("Invalid Excel format. Required columns: 'image name', 'text to voice'")
                return

            self.progress['maximum'] = len(df) + 1
            video_files = []
            api_key_index = 0

            async def process_row(row, index):
                nonlocal api_key_index
                image_name = str(row['image name']).strip()
                text = str(row['text to voice']).strip()
                audio_path = os.path.join(self.output_dir, f"{image_name}.mp3")
                video_path = os.path.join(self.output_dir, f"{image_name}.mp4")

                image_file = None
                for ext in ['.jpg', '.jpeg', '.png']:
                    path = os.path.join(self.image_folder, f"{image_name}{ext}")
                    if os.path.exists(path):
                        image_file = path
                        break
                if not image_file:
                    self.log(f"Image not found for {image_name}")
                    return

                success = False
                while api_key_index < len(self.api_keys) and not success:
                    success = await self.create_voice(text, audio_path, self.api_keys[api_key_index])
                    if not success:
                        api_key_index += 1
                if not success:
                    self.log(f"Failed to create voice for {image_name}")
                    return

                try:
                    probe = ffmpeg.probe(audio_path)
                    duration = float(probe['format']['duration'])
                except ffmpeg.Error as e:
                    self.log(f"Error probing audio {audio_path}: {e.stderr.decode()}")
                    return

                if self.create_segment_video(image_file, audio_path, video_path, duration):
                    video_files.append(video_path)
                self.progress['value'] += 1
                self.root.update()

            loop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)
            tasks = [process_row(row, i) for i, row in df.iterrows()]
            loop.run_until_complete(asyncio.gather(*tasks))
            loop.close()

            if video_files:
                if self.export_dir:
                    random_suffix = uuid.uuid4().hex[:6]
                    output_file = os.path.join(self.export_dir, f"Export-{random_suffix}.mp4")
                    self.combine_videos(video_files, output_file)
                    self.progress['value'] += 1
                    for f in video_files:
                        if os.path.exists(f):
                            os.remove(f)
                    for f in [f for f in os.listdir(self.output_dir) if f.endswith('.mp3')]:
                        os.remove(os.path.join(self.output_dir, f))
                    shutil.rmtree(self.output_dir, ignore_errors=True)
                    self.log("Processing complete!")
                else:
                    self.log("No export folder selected.")
            else:
                self.log("No videos created.")
        except Exception as e:
            self.log(f"Error during processing: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = VideoCreatorApp(root)
    root.mainloop()
