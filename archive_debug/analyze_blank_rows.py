import cv2
import numpy as np
import sys

def count_rows(image_path):
    try:
        # Load image
        img = cv2.imread(image_path)
        if img is None:
            print(f"Error: Could not read image {image_path}")
            return

        # Convert to grayscale
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

        # Edge detection
        edges = cv2.Canny(gray, 50, 150, apertureSize=3)

        # Detect horizontal lines using HoughSheets
        # theta = pi/2 for horizontal lines
        lines = cv2.HoughLinesP(edges, 1, np.pi/180, 100, minLineLength=100, maxLineGap=10)

        horizontal_lines_y = []
        if lines is not None:
            for line in lines:
                x1, y1, x2, y2 = line[0]
                # Filter for horizontal lines (roughly)
                if abs(y1 - y2) < 5: 
                    horizontal_lines_y.append(y1)
        
        # Sort and cluster similar Y values to remove duplicates for thick lines
        horizontal_lines_y.sort()
        unique_lines = []
        if horizontal_lines_y:
            current_group = [horizontal_lines_y[0]]
            for y in horizontal_lines_y[1:]:
                if y - current_group[-1] < 10: # 10 pixel threshold
                    current_group.append(y)
                else:
                    unique_lines.append(int(np.mean(current_group)))
                    current_group = [y]
            unique_lines.append(int(np.mean(current_group)))

        print(f"Found {len(unique_lines)} horizontal lines.")
        # Rows = Lines - 1 (usually)
        print(f"Estimated rows: {len(unique_lines) - 1}")
        
        # Print gaps to see if they are uniform (table rows) or irregular
        gaps = []
        for i in range(len(unique_lines) - 1):
            gaps.append(unique_lines[i+1] - unique_lines[i])
        
        # print(f"Gaps: {gaps}")
        
        # Filter gaps that look like table rows (e.g. median gap)
        if gaps:
            median_gap = np.median(gaps)
            table_rows = [g for g in gaps if 0.8 * median_gap < g < 1.2 * median_gap]
            print(f"Consistent table rows detected: {len(table_rows)}")
            
    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    print("Analyzing EmptyBlank1.jpeg...")
    count_rows("EmptyBlank1.jpeg")
    print("\nAnalyzing EmptyBlank2.jpeg...")
    count_rows("EmptyBlank2.jpeg")
