float4x4 worldViewProj : WorldViewProjection;
 
float cycle : Time < string UIWidget = "none"; >;

texture diffuseTexture : Diffuse
<
	string ResourceName = "4638-v1.bmp";
>;
texture diffuseTexture2 : Diffuse
<
	string ResourceName = "4638-v2.bmp";
>;
texture diffuseTexture3 : Diffuse
<
	string ResourceName = "4638-v3.bmp";
>;
 
float BumpSpeedX <
    string UIName = "Bump Speed X";
    string UIWidget = "slider";
    float UIMin = -0.2;
    float UIMax = 0.2;
    float UIStep = 0.001;
> = -0.01;
float BumpSpeedY <
    string UIName = "Bump Speed Y";
    string UIWidget = "slider";
    float UIMin = -0.2;
    float UIMax = 0.2;
    float UIStep = 0.001;
> = 0.00;
static float2 BumpSpeed = float2(BumpSpeedX,BumpSpeedY);

float TexReptX <
    string UIName = "Texture Repeat X";
    string UIWidget = "slider";
    float UIMin = 1.0;
    float UIMax = 16.0;
    float UIStep = 0.1;
> = 8.0; //2

float TexReptY <
    string UIName = "Texture Repeat Y";
    string UIWidget = "slider";
    float UIMin = 1.0;
    float UIMax = 16.0;
    float UIStep = 0.1;
> = 4.0; //1
static float2 TextureScale = float2(TexReptX,TexReptY);
//------------------------------------
struct vertexInput {
    float4 position				: POSITION;
    float2 texCoordDiffuse		: TEXCOORD0;
};

struct vertexOutput {
    float4 hPosition		: POSITION;
    float2 texCoordDiffuse	: TEXCOORD0;
    float2 texCoordDiffuse2 : TEXCOORD1;
    float2 texCoordDiffuse3 : TEXCOORD2; 
};


//------------------------------------
vertexOutput VS_TransformAndTexture(vertexInput IN) 
{
    vertexOutput OUT;
    OUT.hPosition = mul( float4(IN.position.xyz , 1.0) , worldViewProj);
 
    //float cycle = fmod(Timer, 100.0);	
	OUT.texCoordDiffuse = IN.texCoordDiffuse *TextureScale  + cycle*BumpSpeed ;
    OUT.texCoordDiffuse2 = IN.texCoordDiffuse *TextureScale*2.0 + cycle*BumpSpeed*4.0;
	OUT.texCoordDiffuse3 = IN.texCoordDiffuse *TextureScale*8.0 + cycle*BumpSpeed*16.0;
 
    return OUT;
}


//------------------------------------
sampler TextureSampler = sampler_state 
{
    texture = <diffuseTexture>;
    AddressU  = WRAP;        
    AddressV  = WRAP;
    AddressW  = WRAP;
    MIPFILTER = LINEAR;
    MINFILTER = LINEAR;
    MAGFILTER = LINEAR;
};

sampler TSample2 = sampler_state
{
	texture = <diffuseTexture2>;
    AddressU  = WRAP;        
    AddressV  = WRAP;
    AddressW  = WRAP;
    MIPFILTER = LINEAR;
    MINFILTER = LINEAR;
    MAGFILTER = LINEAR;
};

sampler TSample3 = sampler_state
{
	texture = <diffuseTexture3>;
    AddressU  = WRAP;        
    AddressV  = WRAP;
    AddressW  = WRAP;
    MIPFILTER = LINEAR;
    MINFILTER = LINEAR;
    MAGFILTER = LINEAR;
};

//-----------------------------------
float4 PS_Textured( vertexOutput IN): COLOR
{
  float4 diffuseTexture = tex2D( TextureSampler, IN.texCoordDiffuse );
  float4 dTex2 = tex2D( TSample2, IN.texCoordDiffuse2 );
  float4 dTex3 = tex2D( TSample3, IN.texCoordDiffuse3 );
  return diffuseTexture + (dTex2 * dTex3);
}


//-----------------------------------
technique textured
{
    pass p0 
    {		
    		ZEnable = true;
		ZWriteEnable = true;
		ZFunc = LessEqual;
		CullMode =ccw;
		VertexShader = compile vs_1_1 VS_TransformAndTexture();
		PixelShader  = compile ps_1_1 PS_Textured();
    }
}