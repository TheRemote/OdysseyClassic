#include <ddraw.h>
#include "Windows.h"
#include "stdio.h"
#include <math.h>


long CurTick;
long CurTick2;
char Direction;


static const __int64 Nothing = 0x0101010001010100;


// 16 Bit Masks
static const __int64 RedMask=0xF800F800F800F800;
static const long long GreenMask=0x07E007E007E007E0;
static const long long BlueMask=8725857424048159;		// 001F001F001F001F in decimal


unsigned short *ptrLightMap;
unsigned short *ptrStoredLightMap;

unsigned char *Map;
unsigned char IndoorAlpha;
unsigned char OutdoorAlpha; // Alpha of the light outdoors
unsigned char ShadeAlpha; // This is set by ShadePixel
unsigned char AmbientAlpha; // The ambient alpha for the current lightmap we are drawing

unsigned short LightMap[384][384];
unsigned short StoredLightMap[384][384];

struct Tile {
	unsigned short Ground;
	unsigned short Ground2;
	unsigned short BGTile1;
	unsigned short BGTile2;
	unsigned short FGTile;
	unsigned short FGTile2;
	unsigned char Att;
	unsigned char AttData[4];
	unsigned char Att2;
} Tiles[12][12];




const unsigned char MAX_RAIN_DROPS = 255;
const char RAIN_DROP_SPEED = 6;

const unsigned char MAX_SNOW_FLAKES = 150;
const char SNOW_FLAKE_SPEED = 1;

struct SnowFlakes {
	short x;
	short y;
	char lifetime;
	unsigned char DESTROY;
	char direction;
} SnowFlakes[150];

struct RainDrop {
	short x;
	short y;
	char lifetime;
} RainDrops[255];

unsigned char PixRain[10][3][3] =
{
0, 0, 0,
0, 0, 0,
0, 17, 12,
0, 0, 0,
0, 0 ,0,
0, 24, 18,
0, 0, 0,
0, 13, 9,
0, 15, 11,
0, 0, 0,
0, 27, 19,
0, 22, 15,
0, 28, 19,
0, 26, 18,
0, 0, 0,
0, 26, 18,
0, 35, 19,
0, 0, 0,
0, 21, 13,
0, 42, 25,
0, 0, 0,
0, 44, 25,
0, 0, 0,
0, 0, 0,
0, 51, 28,
0, 0 ,0,
0, 0, 0,
0, 48, 26,
0, 0, 0,
0, 0, 0
};

unsigned char PixSnow[5][5][3] =
{
0, 0, 0,
0, 63, 31,
0, 63, 31,
0, 63, 31,
0, 0, 0,

0, 63, 31,
31, 63, 31,
31, 63, 31,
31, 63, 31,
0, 63, 31,

0, 63, 31,
31, 63, 31,
31, 63, 31,
31, 63, 31,
0, 63, 31,

0, 63, 31,
31, 63, 31,
31, 63, 31,
31, 63, 31,
0, 63, 31,

0, 0, 0,
0, 63, 31,
0, 63, 31,
0, 63, 31,
0, 0, 0
};

struct LightSource 
{
	signed short x;
	signed short y;
	unsigned char Radius;
	unsigned char Intensity;
	unsigned char Permanent;
};

bool ExamineBit(unsigned char Examine, unsigned char Bit)
{
    if (Examine & (int)((float)pow((float)2, (float)Bit)))
		return true;
	else
		return false;
}


void LoadMap()
{
	int position = 0;
	for (int j = 0; j < 12; j++)
	{
		for (int i = 0; i < 12; i++)
		{
			position = 86 + j * 216 + i * 18;
			Tiles[i][j].Ground = Map[position] * 256 + Map[position + 1];
			Tiles[i][j].Ground2 = Map[position + 2] * 256 + Map[position + 3];
			Tiles[i][j].BGTile1 = Map[position + 4] * 256 + Map[position + 5];
			Tiles[i][j].BGTile2 = Map[position + 6] * 256 + Map[position + 7];
			Tiles[i][j].FGTile = Map[position + 8] * 256 + Map[position + 9];
			Tiles[i][j].FGTile2 = Map[position + 10] * 256 + Map[position + 11];
			Tiles[i][j].Att = Map[position + 12];
			Tiles[i][j].AttData[0] = Map[position + 13];
			Tiles[i][j].AttData[1] = Map[position + 14];
			Tiles[i][j].AttData[2] = Map[position + 15];
			Tiles[i][j].AttData[3] = Map[position + 16];
			Tiles[i][j].Att2 = Map[position + 17];
		}
	}
}

static bool ShadePixel(short x, short y, struct LightSource *LS)
{
	static bool Shade;
	static char MapX;
	static char MapY;
	static short FullMapX;
	static short FullMapY;

	Shade = true;
	MapX = x / 32;
	MapY = y / 32;

	switch (Tiles[MapX][MapY].Att)
	{
		case 20: // Light Dampening Tile
			if (Tiles[MapX][MapY].AttData[1] == 0 && Tiles[MapX][MapY].AttData[2] == 0)
			{
				Shade = false;
			}
			else
			{
				FullMapX = MapX * 32;
				FullMapY = MapY * 32;

				if (Tiles[MapX][MapY].AttData[1] > 0)
				{
					if (ExamineBit(Tiles[MapX][MapY].AttData[0], 2) == true)
					{
						if (x < FullMapX + Tiles[MapX][MapY].AttData[1])
						{
							Shade = false;
						}
					}
					else
					{
						if (x > FullMapX + Tiles[MapX][MapY].AttData[1])
						{
							Shade = false;
						}
					}
				}

				if (Tiles[MapX][MapY].AttData[2] > 0)
				{
					if (ExamineBit(Tiles[MapX][MapY].AttData[0], 0) == true)
					{
						if (y < FullMapY + Tiles[MapX][MapY].AttData[2])
						{
							Shade = false;
						}
					}
					else
					{
						if (y > FullMapY + Tiles[MapX][MapY].AttData[2])
						{
							Shade = false;
						}
					}
				}
			}

			if (Shade == false)
			{
				if (ExamineBit(Tiles[MapX][MapY].AttData[3], 1) == true)
				{
					ShadeAlpha = OutdoorAlpha;
				}
				else
				{
					ShadeAlpha = AmbientAlpha;
				}
			}
			break;
		case 22:
			ShadeAlpha = OutdoorAlpha;
			Shade = false;
			break;
	}

	switch (Tiles[MapX][MapY].Att2)
	{
		case 22:
			ShadeAlpha = OutdoorAlpha;
			Shade = false;
			break;
	}


	return Shade;
}

static bool ShadePixel(short x, short y)
{
	static bool Shade;
	static char MapX;
	static char MapY;
	static short FullMapX;
	static short FullMapY;

	Shade = true;
	MapX = x / 32;
	MapY = y / 32;

	switch (Tiles[MapX][MapY].Att)
	{
		case 20:
			if (Tiles[MapX][MapY].AttData[1] == 0 && Tiles[MapX][MapY].AttData[2] == 0)
			{
				Shade = false;
			}
			else
			{
				FullMapX = MapX * 32;
				FullMapY = MapY * 32;

				if (Tiles[MapX][MapY].AttData[1] > 0)
				{
					if (ExamineBit(Tiles[MapX][MapY].AttData[0], 2) == true)
					{
						if (x < FullMapX + Tiles[MapX][MapY].AttData[1])
						{
							Shade = false;
						}
					}
					else
					{
						if (x > FullMapX + Tiles[MapX][MapY].AttData[1])
						{
							Shade = false;
						}
					}
				}

				if (Tiles[MapX][MapY].AttData[2] > 0)
				{
					if (ExamineBit(Tiles[MapX][MapY].AttData[0], 0) == true)
					{
						if (y < FullMapY + Tiles[MapX][MapY].AttData[2])
						{
							Shade = false;
						}
					}
					else
					{
						if (y > FullMapY + Tiles[MapX][MapY].AttData[2])
						{
							Shade = false;
						}
					}
				}
			}

			if (Shade == false)
			{
				if (ExamineBit(Tiles[MapX][MapY].AttData[3], 1) == true)
				{
					ShadeAlpha = OutdoorAlpha;
				}
				else
				{
					ShadeAlpha = AmbientAlpha;
				}
			}
			break;
		case 22:
			ShadeAlpha = OutdoorAlpha;
			Shade = false;
			break;
	}

	switch (Tiles[MapX][MapY].Att2)
	{
		case 22:
			ShadeAlpha = OutdoorAlpha;
			Shade = false;
			break;
	}

	return Shade;
}

long __stdcall ShadeMap16(long sPtr)
{
	try
	{
		__asm
		{
                mov     esi,sPtr			// So we don't have to refer to memory, copy the pointer to the src surface to a register
                mov     edi,ptrLightMap		// Move the pointer to the Light Map here ..

				movq	mm3,RedMask			//load the red colorblend value into the mm3 register
				movq	mm4,GreenMask		//load the green colorblend value into the mm4 register
				movq	mm5,BlueMask		//load the Blue colorblend Value into the mm5 register

                mov     edx,384				// Copy the height of the region to be blitted to the dx register
NextLine:
                mov     ecx,96				// Every time we start a new line, reset the cx register to how many times we need to treat 4 pixels in a line
NextPixels:
				movq	mm7, [edi]			// Get Light Source Values
                movq    mm0, [esi]			// Get 4 source pixel values

                movq    mm2,mm0				// Make a copy of the source pixel value
                pand    mm2,mm3				// Isolate the red values in the source pixels
                psrlw   mm2,11				// Shift the src pixels right to align them to the alpha values
				pmullw  mm2,mm7				// Multiply the src pixels by the alpha value
                psrlw   mm2,8				// Divide by 256 essentially - we have multiplied in total by 255 in the src and dest pixels, so now we've gotta reverse that process
                psllw   mm2,11				// Move the values back to their original bit positions
                movq    mm6,mm2				// Copy the blended red colors to mm6 where the new pixels will be built

                movq    mm2,mm0				// Reset the src pixel value so we get all the colors back

                pand    mm2,mm4				// Isolate the green values in the src pixels
                psrlw   mm2,5				// Shift the src pixels right to align them with the alpha values
                pmullw  mm2,mm7				// Multiply the src pixels by the alpha value
                psrlw   mm2,8				// Divide by 256 - why? explained for the red values
                psllw   mm2,5				// Return the green values to their original bit positions
                por     mm6,mm2				// Copy the blended green colors to mm6 - we're 2/3 of the way there!

                movq    mm2,mm0				// Reset the src pixel value again

                pand    mm2,mm5				// Isolate the blue values in the src pixels
                pmullw  mm2,mm7				// Blue pixels are already aligned with alpha values so we skip to multiplying them by the alpha value
                psrlw   mm2,8				// Divide by 256 - why? explained for the red values
                por     mm6,mm2				// Copy the blended blue colors to mm6 - the pixels have now been completely blended!

                movq    [esi],mm6			// Copy the new pixels to the destination

                add     esi,8				// Set the src pointer to point at the next set of 4 pixels
				add		edi,8

                dec     ecx					// Decrement the width counter
                jnz     NextPixels			// If we've still got more pixels to treat, jump back to the beginning to treat more, otherwise continue on......

                dec     edx					// Decrement the height counter
                jnz     NextLine			// If we've still got lines to treat, jump back to the beginning and get going!  otherwise.......
                emms						// Clean up our mmx registers		
		}
		return 1;
	}
	catch (...)
	{
		return 0;
	}
}

long __stdcall ShadeMap32(long sPtr)
{
	try
	{
		unsigned char *curptr = (unsigned char *)sPtr;
		unsigned short *lightptr = ptrLightMap;
		unsigned short alpha;

		for (int i = 0; i < 147455; i++)
		{
			alpha = *lightptr;

			*curptr = (alpha * (*curptr)) / 256;
			curptr++;

			*curptr = (alpha * (*curptr)) / 256;
			curptr++;

			*curptr = (alpha * (*curptr)) / 256;
			curptr++;

			curptr++;

			lightptr++;
		}

		return 1;
	}
	catch (...)
	{
		return 0;
	}
}

long __stdcall CreateLightMap(struct LightSource *LS, unsigned char ambientalpha, unsigned char *map, unsigned char OutdoorLight)
{
	OutdoorAlpha = 255 - OutdoorLight;
	AmbientAlpha = 255 - ambientalpha;
	Map = map;
	LoadMap();
	try
	{
		static short x, y, c, sx, sy, ex, ey;
		static long d;
		static long r;
		static unsigned short a;
		unsigned short *ptrLS = (unsigned short *)StoredLightMap;

		for (int i = 0; i < 384; i++)
		{
			for (int j = 0; j < 384; j++)
			{
				if (ShadePixel(j, i) == true)
					StoredLightMap[i][j] = AmbientAlpha;
				else
					StoredLightMap[i][j] = ShadeAlpha;
			}
		}

		for(c=0;c<30;c++)
		{
			if(LS->Intensity > 0 && LS->Radius > 0 && LS->Permanent == 1) 
			{
				r = LS->Radius * LS->Radius;
				if (LS->x - LS->Radius < 0) {sx=0;} else { sx = LS->x - LS->Radius; }
				if (LS->y - LS->Radius < 0) {sy=0;} else { sy = LS->y - LS->Radius; }
				if (LS->x + LS->Radius > 383) {ex=383;} else { ex = LS->x + LS->Radius; }
				if (LS->y + LS->Radius > 383) {ey=383;} else { ey = LS->y + LS->Radius; }
				for(y=sy;y<=ey;y++) 
				{
					for(x=sx;x<=ex;x++) 
					{
						d = (x-LS->x)*(x-LS->x) + (y-LS->y)*(y-LS->y);
						if (d <= r) 
						{
							if (ShadePixel(x, y, LS) == true)
							{
								a = (unsigned short)(d * (LS->Intensity)/r);
								a = (LS->Intensity) - a;
								a + StoredLightMap[y][x] > 255 ? StoredLightMap[y][x] = 255 : StoredLightMap[y][x] += a;
							}
							else
							{
								StoredLightMap[y][x] = ShadeAlpha;
							}
						}
					}
				}
			}
			LS++;
		}
		return 0;
	}
	catch (...)
	{
		return 1;
	}
}

long __stdcall UpdateLightMap(struct LightSource *LS)
{
	try
	{
		static short x, y, c, sx, sy, ex, ey;
		static long d;
		static long r;
		static unsigned short a;
		unsigned short *ptrLS = (unsigned short *)LightMap;

		memcpy(LightMap, StoredLightMap, sizeof(LightMap));

		for(c=0;c<30;c++)
		{
			if(LS->Intensity > 0 && LS->Radius > 0 && LS->Permanent == 0) 
			{
				r = LS->Radius * LS->Radius;
				if (LS->x - LS->Radius < 0) {sx=0;} else { sx = LS->x - LS->Radius; }
				if (LS->y - LS->Radius < 0) {sy=0;} else { sy = LS->y - LS->Radius; }
				if (LS->x + LS->Radius > 383) {ex=383;} else { ex = LS->x + LS->Radius; }
				if (LS->y + LS->Radius > 383) {ey=383;} else { ey = LS->y + LS->Radius; }
				for(y=sy;y<=ey;y++) 
				{
					for(x=sx;x<=ex;x++) 
					{
						d = (x-LS->x)*(x-LS->x) + (y-LS->y)*(y-LS->y);
						if (d <= r) 
						{
							if (ShadePixel(x, y, LS) == true)
							{
								a = (unsigned short)(d * (LS->Intensity)/r);
								a = (LS->Intensity) - a;
								a + LightMap[y][x] > 255 ? LightMap[y][x] = 255 : LightMap[y][x] += a;
							}
							else
							{
								LightMap[y][x] = ShadeAlpha;
							}
						}
					}
				}
			}
			LS++;
		}
		return 0;
	}
	catch (...)
	{
		return 1;
	}
}


long __stdcall InitializeLighting()
{
	ptrLightMap = (unsigned short *)LightMap;
	return 0;
}

long __stdcall InitRain(long TickCount)
{
	short A;
	srand((unsigned long)TickCount);
	for(A=0;A<MAX_RAIN_DROPS;A++)
	{
		RainDrops[A].lifetime = 12 + (rand()>>10);
		RainDrops[A].x = (rand()/86);
		RainDrops[A].y = (rand()/84)-(RainDrops[A].lifetime * RAIN_DROP_SPEED);
	}
	CurTick = 0;
	return 0;
}

void ReInitRainDrop(unsigned char CurDrop) 
{
	RainDrops[CurDrop].lifetime = 12 + (rand()>>11);
	RainDrops[CurDrop].x = (rand()/86);
	RainDrops[CurDrop].y = (rand()/84)-(RainDrops[CurDrop].lifetime * RAIN_DROP_SPEED);
}

long __stdcall Rain16(long sfcptr, long TickCount)
{
	try
	{
		static unsigned char A, B, C, D, F, G;
		static unsigned short E;
		unsigned short *curptr = (unsigned short *)sfcptr;

		for(A=0;A<MAX_RAIN_DROPS;A++) 
		{
			if (TickCount >= CurTick) 
			{
				if (Direction == 1) 
				{
					RainDrops[A].x += RAIN_DROP_SPEED>>2;
				} 
				else if (Direction ==2) 
				{
					RainDrops[A].x -= RAIN_DROP_SPEED>>2;
				}
				B = 1;
				RainDrops[A].y += (RAIN_DROP_SPEED + rand() % 4);
				RainDrops[A].lifetime--;
				if (RainDrops[A].y > 384 || RainDrops[A].lifetime <= 0) 
				{
					ReInitRainDrop(A);
				}
			}
			for(C=0;C<10;C++) 
			{
				for(D=0;D<3;D++) 
				{
					if (Direction ==1) 
					{
						F = C;
						G = abs(D - 2);
					} 
					else 
					{
						F = C;
						G = D;
					}
					if (RainDrops[A].y + (short)C > 0 && RainDrops[A].y + (short)C < 384) 
					{
						if (!(PixRain[F][G][0] == 0 && PixRain[F][G][1] == 0 && PixRain[F][G][2] == 0)) 
						{
							curptr = (unsigned short *)sfcptr + ((RainDrops[A].y + C) * 384 + RainDrops[A].x + D);
							*curptr = PixRain[F][G][0] << 11 | PixRain[F][G][1] << 5 | PixRain[F][G][2];
						}
					}
				}
			}
		}
		if (B == 1) 
		{
			CurTick = TickCount + 50;
			B = 0;
		}
		return CurTick;
	}
	catch (...)
	{
		return 0;
	}
}


long __stdcall InitSnow(long TickCount)
{
	short A;
	srand((unsigned long)TickCount);
	for(A=0;A<MAX_SNOW_FLAKES;A++)
	{
		SnowFlakes[A].lifetime = 20 + (rand()>>10);
		SnowFlakes[A].x = (rand()/86);
		SnowFlakes[A].y = (rand()/84)-(SnowFlakes[A].lifetime * SNOW_FLAKE_SPEED);
		if (rand()%2 == 1)
			SnowFlakes[A].DESTROY = (unsigned char)((short)(((double)rand()/32767) * 255));
		else
			SnowFlakes[A].DESTROY = 255;
	}
	CurTick = 0;
	return 0;
}

void ReInitSnowFlake(unsigned short CurDrop) {
	SnowFlakes[CurDrop].lifetime = 20 + (rand()>>10);
	SnowFlakes[CurDrop].x = (rand()/86);
	SnowFlakes[CurDrop].y = (rand()/84)-(SnowFlakes[CurDrop].lifetime * SNOW_FLAKE_SPEED );
	SnowFlakes[CurDrop].DESTROY = 255;
}

long __stdcall Snow16(long sfcptr, long TickCount)
{
	try
	{
		static unsigned short A, B, C, D;
		unsigned short *curptr = (unsigned short *)sfcptr;

		for(A=0;A<MAX_SNOW_FLAKES;A++) 
		{
			if (TickCount >= CurTick) 
			{
				B = 1;
				if (SnowFlakes[A].DESTROY == 255) 
				{
					SnowFlakes[A].y += SNOW_FLAKE_SPEED;
					C = rand();
					if (TickCount >= CurTick2) 
					{ 
						if (C < 10000) 
						{
							//SnowFlakes[A].x--;
							SnowFlakes[A].direction = 2;
							SnowFlakes[A].x--;
						} 
						else if (C > 22767) 
						{
							//SnowFlakes[A].x++;
							SnowFlakes[A].direction = 1;
							SnowFlakes[A].x++;
						}
					
						CurTick2 = TickCount + 100;
					}
					if (C > 10000) 
					{ 
						if (C < 22767) 
						{
							if (SnowFlakes[A].direction == 1) 
							{
								SnowFlakes[A].x++;	
							} 
							else if (SnowFlakes[A].direction ==2) 
							{
								SnowFlakes[A].x--;
							}

						}
					}
					if (SnowFlakes[A].direction == 0) 
					{
						CurTick2 = CurTick;
					}
				}
				SnowFlakes[A].lifetime--;
				if (SnowFlakes[A].y > 384) 
				{
					ReInitSnowFlake(A);
				}
				if (SnowFlakes[A].lifetime == 0) 
				{
					if (SnowFlakes[A].DESTROY >= 5) 
					{
						SnowFlakes[A].DESTROY -=5;
						SnowFlakes[A].lifetime = 1;
					} 
					else 
					{
						ReInitSnowFlake(A);
					}
				}
			}
			for(C=0;C<5;C++) 
			{
				for(D=0;D<5;D++) 
				{
					if (SnowFlakes[A].y + (short)C > 0 && SnowFlakes[A].y + (short)C < 384) 
					{
						if (SnowFlakes[A].x + (short)C > 0 && SnowFlakes[A].x + (short)C < 384) 
						{
							if (!(PixSnow[C][D][0] == 0 && PixSnow[C][D][1] == 0 && PixSnow[C][D][2] == 0)) 
							{
								curptr = (unsigned short *)(sfcptr + ((SnowFlakes[A].y + C) * 384 + SnowFlakes[A].x + D));
								//*curptr = PixSnow[C][D][0] << 11 | PixSnow[C][D][1] << 5 | PixSnow[C][D][2];
								*curptr = (63488 & ((*curptr & 63488) + ((((PixSnow[C][D][0]<<11) - (*curptr & 63488)) * SnowFlakes[A].DESTROY) >>8))) | (2016 & ((*curptr & 2016) + ((((PixSnow[C][D][1]<<5) - (*curptr & 2016)) * SnowFlakes[A].DESTROY) >>8))) | (31 & ((*curptr & 31) + (((PixSnow[C][D][2] - (*curptr & 31)) * SnowFlakes[A].DESTROY) >>8)));
							}
						}
					}
				}
			}
		}
		if (B == 1) 
		{
			CurTick = TickCount + 50;
			B = 0;
		}
		return CurTick;
	}
	catch (...)
	{
		return 0;
	}
}

long __stdcall Rain32(long sfcptr, long TickCount)
{
	try
	{
		static unsigned char A, B, C, D, F, G;
		static unsigned short E;
		static unsigned long oldcolor;
		static unsigned short pixel;
		unsigned char *curptr = (unsigned char *)sfcptr;

		for(A=0;A<MAX_RAIN_DROPS;A++) 
		{
			if (TickCount >= CurTick) 
			{
				if (Direction == 1) 
				{
					RainDrops[A].x += RAIN_DROP_SPEED>>2;
				} 
				else if (Direction ==2) 
				{
					RainDrops[A].x -= RAIN_DROP_SPEED>>2;
				}
				B = 1;
				RainDrops[A].y += (RAIN_DROP_SPEED + rand() % 4);
				RainDrops[A].lifetime--;
				if (RainDrops[A].y > 383 || RainDrops[A].lifetime <= 0) 
				{
					ReInitRainDrop(A);
				}
			}
			for(C=0;C<10;C++) 
			{
				for(D=0;D<3;D++) 
				{
					if (Direction ==1) 
					{
						F = C;
						G = abs(D - 2);
					} 
					else 
					{
						F = C;
						G = D;
					}
					if (RainDrops[A].y + (short)C > 0 && RainDrops[A].y + (short)C < 384) 
					{
						if (!(PixRain[F][G][0] == 0 && PixRain[F][G][1] == 0 && PixRain[F][G][2] == 0)) 
						{
							curptr = (unsigned char *)(sfcptr + (((RainDrops[A].y + C) * 384 + RainDrops[A].x + D)*4));

							pixel = PixRain[F][G][0] << 11 | PixRain[F][G][1] << 5 | PixRain[F][G][2];

							oldcolor = *curptr;
							*curptr = ((((pixel & 0x001F)) * 0xFF) / 0x1F);
							curptr++;
							oldcolor = *curptr;
							*curptr = ((((pixel & 0x07E0) >> 5) * 0xFF) / 0x3F);
							curptr++;
							oldcolor = *curptr;
							*curptr = ((((pixel & 0xF800) >> 11) * 0xFF) / 0x1F);
						}
					}
				}
			}
		}
		if (B == 1) 
		{
			CurTick = TickCount + 50;
			B = 0;
		}
		return CurTick;
	}
	catch (...)
	{
		return 0;
	}
}

long __stdcall Snow32(long sfcptr, long TickCount)
{
	try
	{
		static unsigned short A, B, C, D;
		static unsigned short pixel;
		static unsigned char alpha;
		static unsigned long oldcolor;
		unsigned char *curptr = (unsigned char *)sfcptr;

		for(A=0;A<MAX_SNOW_FLAKES;A++) 
		{
			if (TickCount >= CurTick) 
			{
				B = 1;
				if (SnowFlakes[A].DESTROY == 255) 
				{
					SnowFlakes[A].y += SNOW_FLAKE_SPEED;
					C = rand();
					if (TickCount >= CurTick2) 
					{ 
						if (C < 10000) 
						{
							//SnowFlakes[A].x--;
							SnowFlakes[A].direction = 2;
							SnowFlakes[A].x--;
						} 
						else if (C > 22767) 
						{
							//SnowFlakes[A].x++;
							SnowFlakes[A].direction = 1;
							SnowFlakes[A].x++;
						}
					
						CurTick2 = TickCount + 100;
					}
					if (C > 10000) 
					{ 
						if (C < 22767) 
						{
							if (SnowFlakes[A].direction == 1) 
							{
								SnowFlakes[A].x++;	
							} 
							else if (SnowFlakes[A].direction ==2) 
							{
								SnowFlakes[A].x--;
							}

						}
					}
					if (SnowFlakes[A].direction == 0) 
					{
						CurTick2 = CurTick;
					}
				}
				SnowFlakes[A].lifetime--;
				if (SnowFlakes[A].y > 383) 
				{
					ReInitSnowFlake(A);
				}
				if (SnowFlakes[A].lifetime == 0) 
				{
					if (SnowFlakes[A].DESTROY >= 5) 
					{
						SnowFlakes[A].DESTROY -=5;
						SnowFlakes[A].lifetime = 1;
					} 
					else 
					{
						ReInitSnowFlake(A);
					}
				}
			}
			for(C=0;C<5;C++) 
			{
				for(D=0;D<5;D++) 
				{
					if (SnowFlakes[A].y + (short)C > 0 && SnowFlakes[A].y + (short)C < 384) 
					{
						if (SnowFlakes[A].x + (short)C > 0 && SnowFlakes[A].x + (short)C < 384) 
						{
							if (!(PixSnow[C][D][0] == 0 && PixSnow[C][D][1] == 0 && PixSnow[C][D][2] == 0)) 
							{
								alpha = 255 - SnowFlakes[A].DESTROY;
								curptr = (unsigned char *)(sfcptr + (((SnowFlakes[A].y + C) * 384 + SnowFlakes[A].x + D)*4));

								pixel = PixSnow[C][D][0] << 11 | PixSnow[C][D][1] << 5 | PixSnow[C][D][2];

								oldcolor = *curptr;
								*curptr = ((((pixel & 0x001F)) * 0xFF) / 0x1F);
								*curptr = (unsigned char)(( alpha * ( oldcolor - *curptr ) ) / 256 + *curptr);
								curptr++;
								oldcolor = *curptr;
								*curptr = ((((pixel & 0x07E0) >> 5) * 0xFF) / 0x3F);
								*curptr = (unsigned char)(( alpha * ( oldcolor - *curptr ) ) / 256 + *curptr);
								curptr++;
								oldcolor = *curptr;
								*curptr = ((((pixel & 0xF800) >> 11) * 0xFF) / 0x1F);
								*curptr = (unsigned char)(( alpha * ( oldcolor - *curptr ) ) / 256 + *curptr);
							}
						}
					}
				}
			}
		}
		if (B == 1) 
		{
			CurTick = TickCount + 50;
			B = 0;
		}
		return CurTick;
	}
	catch (...)
	{
		return 0;
	}
}

__int32 _stdcall EncryptDataFile(char *strptr, unsigned char XOR)
{
	char FileName[512];
	char Buffer[32768];
	char *ptrBuffer = Buffer;
	FILE *CurFile;
	__int32 FileSize, ReadLength, cs;
	__int64 XORValue;
	
	memset(Buffer, 0, sizeof(Buffer));
	memset(FileName, 0, sizeof(FileName));
	strcpy(FileName, strptr);
	
	XORValue = (XOR) + (XOR << 8) + (XOR << 16) + (XOR << 24) + (XOR << 32) + (XOR << 40) + (XOR << 48) + (XOR << 56);

	CurFile = fopen(FileName,"rb+");
	if(CurFile == NULL) 
	{
		return 1;
	} 
	else 
	{
		fseek(CurFile,0,SEEK_END);
		FileSize = ftell(CurFile);
		rewind(CurFile);
		while(FileSize)
		{
			if (FileSize < 32768) 
			{
				ReadLength = FileSize;
				FileSize = 0;
			} else 
			{
				ReadLength = 32768;
				FileSize -= 32768;
			}
			cs = ftell(CurFile);
			fread(Buffer,1,ReadLength,CurFile);
			__asm
			{
				mov     esi,4096			// 4096 iterations
				mov		edi,ptrBuffer
				movq	mm1,XORValue
				NextQuadWord:
				movq    mm0,[edi]		// move 8 bytes into the mm0 register
				pxor	mm0,mm1
				movq	[edi],mm0
				add		edi,8
				dec		esi
				jnz		NextQuadWord
				emms
			}
			fseek(CurFile,cs,SEEK_SET);
			fwrite(Buffer,1,ReadLength,CurFile);
		}
	}
	fclose(CurFile);
	return 0;
}

__int32 _stdcall EncryptDataString(char *strptr, unsigned char XOR)
{
	try
	{
		char Buffer[4096];
		char *ptrBuffer = Buffer;
		__int64 XORValue;
		memset(Buffer, 0, sizeof(Buffer));
		strcpy(Buffer, strptr);
		
		XORValue = (XOR) + (XOR << 8) + (XOR << 16) + (XOR << 24) + (XOR << 32) + (XOR << 40) + (XOR << 48) + (XOR << 56);

		__asm
		{
			mov     esi,512			// 512 iterations
			mov		edi,ptrBuffer
			movq	mm1,XORValue
			NextQuadWord:
			movq    mm0,[edi]		// move 8 bytes into the mm0 register
			pxor	mm0,mm1
			movq	[edi],mm0
			add		edi,8
			dec		esi
			jnz		NextQuadWord
			emms
		}

		strcpy(strptr, Buffer);
	}
	catch (...)
	{

	}
	
	return 0;
}