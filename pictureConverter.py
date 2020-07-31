from PIL import Image

image = Image.open('monalisa.jpg')

with open('pixels.csv', 'a') as outputFile:
	for y in range(0, image.size[1]):
		line = ''
		for x in range(0, image.size[0]):
			for colorValue in image.getpixel((x, y)):
				line = line + str(colorValue) + ' '
			line = line[:-1] + ','
		outputFile.write(line[:-1] + '\n')
