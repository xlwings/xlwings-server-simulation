import numpy as np
import xlwings as xw
from fastapi import APIRouter, Body, Security

from ..core.auth import authenticate

# Require authentication for all endpoints for this router
router = APIRouter(
    dependencies=[Security(authenticate)],
    prefix="/simulation",
    tags=["Simulation"],
)


@router.post("/monte-carlo")
async def simulation(data: dict = Body):
    book = xw.Book(json=data)
    sht = book.sheets[0]

    # User Inputs
    num_simulations = sht.range('E3').options(numbers=int).value
    time = sht.range('E4').value
    num_timesteps = sht.range('E5').options(numbers=int).value
    dt = time/num_timesteps  # Length of time period
    vol = sht.range('E7').value
    mu = np.log(1 + sht.range('E6').value)  # Drift
    starting_price = sht.range('E8').value
    perc_selection = [5, 50, 95]  # percentiles

    # Excel: clear output, write out initial values of percentiles/sample path and set chart source
    # and x-axis values
    sht.range('O2').expand().clear_contents()
    sht.range('P2').value = [starting_price, starting_price, starting_price, starting_price]
    # Not yet implemented with remote interpreter
    # sht.charts['Chart 5'].set_source_data(sht.range((1, 15), (num_timesteps + 2, 19)))
    sht.range('O2').value = np.round(np.linspace(0, time, num_timesteps + 1).reshape(-1, 1), 2)

    # Preallocation
    price = np.zeros((num_timesteps + 1, num_simulations))
    percentiles = np.zeros((num_timesteps + 1, 3))

    # Set initial values
    price[0, :] = starting_price
    percentiles[0, :] = starting_price

    # Simulation at each time step
    for t in range(1, num_timesteps + 1):
        rand_nums = np.random.randn(num_simulations)
        price[t,:] = price[t-1, :] * np.exp((mu - 0.5 * vol**2) * dt + vol * rand_nums * np.sqrt(dt))
        percentiles[t, :] = np.percentile(price[t, :], perc_selection)

    sht.range('P2').value = percentiles
    sht.range('S2').value = price[:, :1]  # Sample path

    return book.json()
