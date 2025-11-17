import os
import discord
from discord.ext import commands
from dotenv import load_dotenv
import excel_work

load_dotenv()
TOKEN = os.getenv("DISCORD_BOT_TOKEN")

intents = discord.Intents.default()  # no message_content needed for slash
bot = commands.Bot(command_prefix="!", intents=intents)  # prefix won't be used


@bot.event
async def on_ready():
    # sync slash commands (global sync may take ~1h; per-guild is instant)
    await bot.tree.sync()
    print(f"✅ Logged in as {bot.user} (id: {bot.user.id})")


# Example slash command
@bot.tree.command(name="hello", description="Say hello")
async def hello(interaction: discord.Interaction):
    await interaction.response.send_message(f"Hello, {interaction.user.mention}! 👋")


@bot.tree.command(name="add", description="Add two numbers")
async def add(interaction: discord.Interaction, x: int, y: int):
    result = x + y
    await interaction.response.send_message(f"{x} + {y} = {result}")


@bot.tree.command(name="dm_me", description="Send yourself a DM")
async def dm_me(interaction: discord.Interaction, message: str):
    try:
        await interaction.user.send(f"Here’s your DM: {message}")
        await interaction.response.send_message("✅ I sent you a DM!", ephemeral=True)
    except discord.Forbidden:
        await interaction.response.send_message("❌ I can’t DM you (maybe you have DMs disabled).", ephemeral=True)


@bot.event
async def on_message(message: discord.Message):
    if message.author.bot:
        return
    if isinstance(message.channel, discord.DMChannel):
        await message.channel.send(f"Hi {message.author.name}, I got your DM: {message.content}")

@bot.tree.command(name="annoy", description="Annoy her")
async def fact(interaction: discord.Interaction):
    user2_id = 707366300267315243
    await interaction.response.send_message(f'Heyyy bitchhh, <@{user2_id}>. \n how ya doing??')


#TODO Command to get the final price of who owes who, pinging the person
@bot.tree.command(name="expenses", description="Gets who owes who money and how much")
async def expenses(interaction: discord.Interaction, month: str):
    # Ping multiple users by ID
    user1_id = 141180424964669440
    user2_id = 707366300267315243

    print('Calling function to get expenses result')
    await interaction.response.defer()
    try:
        result = excel_work.main_function(month)
        print(result)
        await interaction.followup.send(f'<@{user1_id}> <@{user2_id}> {result}')
    except Exception as e:
        await interaction.followup.send(f"There was the following error: {e}")


#TODO Read for command to write on excel file that the months are payed



 
bot.run(TOKEN)
